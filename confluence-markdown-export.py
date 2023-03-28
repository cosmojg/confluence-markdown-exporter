import argparse
import os
from time import sleep
from urllib.parse import urlparse, urlunparse

import doc2docx
import pypandoc
import requests
from atlassian import Confluence

DOWNLOAD_CHUNK_SIZE = (
    4 * 1024 * 1024
)  # 4MB, since we're single threaded this is safe to raise much higher


class ExportException(Exception):
    pass


class Exporter:
    def __init__(self, url, username, token, out_dir, space) -> None:
        self.__out_dir = out_dir
        self.__parsed_url = urlparse(url)
        self.__username = username
        self.__token = token
        self.__confluence = Confluence(
            url=urlunparse(self.__parsed_url),
            username=self.__username,
            password=self.__token,
        )
        self.__seen = set()
        self.__space = space

    def __sanitize_filename(self, document_name_raw):
        document_name = document_name_raw
        document_name = document_name.replace(".doc", "xdoc")
        for invalid in ["..", ".", "/"]:
            if invalid in document_name:
                print(
                    (
                        f'Dangerous page title: "{document_name}", "{invalid}" found,'
                        ' replacing it with "_"'
                    ),
                )
                document_name = document_name.replace(invalid, "_")
        document_name = document_name.replace("xdoc", ".doc")
        return document_name

    def __dump_page(self, src_id, parents):
        if src_id in self.__seen:
            # this could theoretically happen if Page IDs are not unique or there is a
            # circle
            raise ExportException("Duplicate Page ID Found!")

        page = self.__confluence.get_page_by_id(src_id, expand="body.storage")
        page_title = page["title"]
        page_id = page["id"]

        # see if there are any children
        child_ids = self.__confluence.get_child_id_list(page_id)

        # save all files as .doc for now, we will convert them later
        document_name = "home" + ".doc" if len(child_ids) > 0 else f"{page_title}.doc"

        # make some rudimentary checks, to prevent trivial errors
        sanitized_filename = self.__sanitize_filename(document_name)
        sanitized_parents = list(map(self.__sanitize_filename, parents))

        page_location = [*sanitized_parents, sanitized_filename]
        page_filename_doc = os.path.join(self.__out_dir, *page_location)
        page_filename_docx = page_filename_doc.replace(".doc", ".docx")
        page_filename_md = page_filename_doc.replace(".doc", ".md")
        page_filename_media = page_filename_doc.replace(".doc", "_media")

        page_output_dir = os.path.dirname(page_filename_doc)
        os.makedirs(page_output_dir, exist_ok=True)

        print(f"Saving to {' / '.join(page_location)}")
        r = requests.get(
            f"https://confluence.bostonfusion.com/exportword?pageId={page_id}",
            auth=(self.__username, self.__token),
            stream=True,
        )

        print(f"Writing {page_filename_doc}")
        with open(page_filename_doc, "wb") as f:
            f.write(r.content)

        print(f"Writing {page_filename_docx}")
        doc2docx.convert(page_filename_doc, page_filename_docx)

        attempt = 1
        wait = 2
        for _i in range(4):
            try:
                print(f"Writing {page_filename_md} (Attempt #{attempt})")
                pypandoc.convert_file(
                    page_filename_docx,
                    "gfm",
                    outputfile=page_filename_md,
                    extra_args=["--extract-media", page_filename_media],
                )
                error = None
            except Exception as e:
                error = str(e)

            if not error:
                break

            print(error)
            print(f"Waiting {wait} seconds before trying again...")
            sleep(wait)
            wait *= 2
            attempt += 1
        self.__seen.add(page_id)

        # recurse to process child nodes
        for child_id in child_ids:
            self.__dump_page(child_id, parents=[*sanitized_parents, page_title])

    def __dump_space(self, space):
        space_key = space["key"]
        print("Processing space", space_key)
        if space.get("homepage") is None:
            print(f"Skipping space: {space_key}, no homepage found!")
            print("In order for this tool to work there has to be a root page!")
        else:
            # homepage found, recurse from there
            homepage_id = space["homepage"]["id"]
            self.__dump_page(homepage_id, parents=[space_key])

    def dump(self):
        ret = self.__confluence.get_all_spaces(
            start=0,
            limit=500,
            expand="description.plain,homepage",
        )
        if ret["size"] == 0:
            print("No spaces found in confluence. Please check credentials")
        for space in ret["results"]:
            if self.__space is None or space["key"] == self.__space:
                self.__dump_space(space)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("url", type=str, help="link to Confluence instance")
    parser.add_argument("username", type=str, help="username")
    parser.add_argument("token", type=str, help="personal access token or password")
    parser.add_argument("out_dir", type=str, help="output directory")
    parser.add_argument(
        "--space",
        type=str,
        required=False,
        default=None,
        help="space(s) to export",
    )
    args = parser.parse_args()

    dumper = Exporter(
        url=args.url,
        username=args.username,
        token=args.token,
        out_dir=args.out_dir,
        space=args.space,
    )
    dumper.dump()
