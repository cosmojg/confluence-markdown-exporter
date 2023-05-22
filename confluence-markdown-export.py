import argparse
import os
import re
import shutil
import sys
import tempfile
from pathlib import Path
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

    def __sanitize(self, page_title):
        page_title = re.sub("[\\\\/\\[\\] ]+", "-", page_title)
        page_title = re.sub("\\s+", "-", page_title)
        page_title = re.sub("\\.\\.+", ".", page_title)
        page_title = re.sub("--+", "-", page_title)
        page_title = re.sub("-(-|\\s)+-", "-", page_title)
        page_title = page_title.strip("-. ")
        return page_title

    def __download(self, page_id, page_filename_doc):
        print(f"Generating {page_filename_doc}")
        r = requests.get(
            f"https://confluence.bostonfusion.com/exportword?pageId={page_id}",
            auth=(self.__username, self.__token),
            stream=True,
        )
        with open(page_filename_doc, "wb") as f:
            f.write(r.content)

    def __modernize(self, page_filename_doc, page_filename_docx):
        print(f"Generating {page_filename_docx}")
        if sys.platform == "darwin":
            tempdir = tempfile.TemporaryDirectory(
                dir=f"{Path.home()}/Library/Containers/com.microsoft.Word/Data/Documents",
            )
        elif sys.platform == "win32":
            tempdir = tempfile.TemporaryDirectory(dir=self.__out_dir)
        else:
            msg = "Incompatible operating system (Microsoft Word must be installed)"
            raise NotImplementedError(
                msg,
            )
        tempdoc = os.path.join(tempdir.name, "x.doc")
        tempdocx = tempdoc.replace(".doc", ".docx")
        shutil.copyfile(page_filename_doc, tempdoc)
        doc2docx.convert(tempdoc, tempdocx)
        try:
            shutil.copyfile(tempdocx, page_filename_docx)
        except Exception:
            tempdir.cleanup()
            raise
        tempdir.cleanup()

    def __convert(self, page_filename_docx, page_filename_md, page_filename_media):
        print(f"Generating {page_filename_md}")
        pypandoc.convert_file(
            page_filename_docx,
            "gfm",
            outputfile=page_filename_md,
            extra_args=["--extract-media", page_filename_media],
        )
        md_path = Path(page_filename_md)
        md_path.write_text(md_path.read_text().replace(f"{md_path.parent!s}/", ""))

    def __dump_page(self, src_id, parents, home_md):
        if src_id in self.__seen:
            # this could theoretically happen if Page IDs are not unique or there is a circle
            msg = "Duplicate Page ID Found!"
            raise ExportException(msg)

        page = self.__confluence.get_page_by_id(src_id, expand="body.storage")
        page_title = page["title"]
        page_id = page["id"]

        # see if there are any children
        child_ids = self.__confluence.get_child_id_list(page_id)

        # save all files as .doc for now, we will convert them later
        # make some rudimentary checks, to prevent trivial errors
        sanitized_parents = list(map(self.__sanitize, parents))
        page_location = [*sanitized_parents, self.__sanitize(page_title)]
        page_filename_doc = f"{os.path.join(self.__out_dir, *page_location)}.doc"
        page_filename_docx = page_filename_doc.replace(".doc", ".docx")
        page_filename_md = page_filename_doc.replace(".doc", ".md")
        page_filename_media = page_filename_doc.removesuffix(".doc")

        page_output_dir = os.path.dirname(page_filename_doc)
        os.makedirs(page_output_dir, exist_ok=True)
        home_md.write_text(
            f"{home_md.read_text()}\n* [{page_filename_media}]({page_filename_media})",
        )
        home_md.write_text(home_md.read_text().replace(f"{home_md.parent!s}/", ""))

        # Download .doc from Confluence
        if not Path(page_filename_doc).is_file():
            try:
                self.__download(page_id, page_filename_doc)
            except Exception as e:
                print(e)
                print("Waiting 10 seconds before trying again...")
                sleep(10)
                self.__download(page_id, page_filename_doc)

        # Convert .doc to .docx
        if not Path(page_filename_docx).is_file():
            try:
                self.__modernize(page_filename_doc, page_filename_docx)
            except Exception as e:
                print(e)
                print("Waiting 10 seconds before trying again...")
                sleep(10)
                self.__download(page_id, page_filename_doc)
                self.__modernize(page_filename_doc, page_filename_docx)

        # Attempt to convert .docx to .md
        if not Path(page_filename_md).is_file():
            try:
                self.__convert(
                    page_filename_docx,
                    page_filename_md,
                    page_filename_media,
                )
            except Exception as e:
                print(e)
                print("Waiting 10 seconds before trying again...")
                sleep(10)
                self.__download(page_id, page_filename_doc)
                self.__modernize(page_filename_doc, page_filename_docx)
                self.__convert(
                    page_filename_docx,
                    page_filename_md,
                    page_filename_media,
                )

        # Mark page as seen
        self.__seen.add(page_id)

        # recurse to process child nodes
        for child_id in child_ids:
            self.__dump_page(
                child_id,
                parents=[*sanitized_parents, page_title],
                home_md=home_md,
            )

    def __dump_space(self, space):
        space_key = space["key"]
        print("Processing space", space_key)
        if space.get("homepage") is None:
            print(f"Skipping space: {space_key}, no homepage found!")
            print("In order for this tool to work there has to be a root page!")
        else:
            # homepage found, recurse from there
            homepage_id = space["homepage"]["id"]
            sanitized_parents = list(map(self.__sanitize, [space_key]))
            page_location = [*sanitized_parents, "home"]
            home_md = Path(f"{os.path.join(self.__out_dir, *page_location)}.md")
            os.makedirs(str(home_md.parent), exist_ok=True)
            home_md.write_text("# Migrated from Confluence:")
            self.__dump_page(homepage_id, parents=[space_key], home_md=home_md)

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
