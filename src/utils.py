import logging
import re
from json import loads
from os.path import basename

from rich.logging import RichHandler
from rich.progress import Progress

logger = logging.getLogger("ZoteroCitation")
formatter = logging.Formatter("%(name)s :: %(message)s", datefmt="%m-%d %H:%M:%S")
handler = RichHandler()
handler.setFormatter(formatter)
logger.addHandler(handler)
logger.setLevel(logging.INFO)


def get_citations_info(docx_obj) -> dict[str, dict[str, str]]:
    """
    Scan the word, collect title, publisher, author and itemKey for each citation.

    :param docx_obj: Opened word get from pywin32.
    :return: Collected title (key) and itemKey (value).
    """
    titles_item_keys = {}
    total = len(list(docx_obj.Fields))

    with Progress() as progress:
        pid = progress.add_task(f"[red]Collecting titles, publishers, authors and itemKeys for citations..[red]", total=total)

        for field in docx_obj.Fields:
            progress.advance(pid, advance=1)

            if "ADDIN ZOTERO_ITEM" in field.Code.Text:
                # convert string to JSON string.
                field_value: str = field.Code.Text.strip()
                field_value = field_value.strip("ADDIN ZOTERO_ITEM CSL_CITATION").strip()
                field_value_json = loads(field_value)
                citations_list = field_value_json["citationItems"]

                for _citation in citations_list:
                    # pprint(_citation)
                    # raise ZoteroCitationError
                    item_key = basename(_citation["uris"][0])
                    title = _citation["itemData"]["title"]

                    if "publisher" in _citation["itemData"]:
                        publisher = _citation["itemData"]["publisher"]
                    else:
                        publisher = ""

                    author = _citation["itemData"]["author"][0]
                    if "family" in author:
                        author = author["family"]
                    else:
                        author = author["literal"]

                    if title not in titles_item_keys:
                        titles_item_keys[title] = {
                            "item_key": item_key,
                            "publisher": publisher,
                            "author": author,
                        }

    return titles_item_keys


def replace_invalid_char(text: str) -> str:
    """
    Replace invalid characters with "" because bookmarks in Word mustn't contain these characters.

    :param text: Input text.
    :type text: str
    :return: Text in which all invalid characters have been replaced.
    :rtype: str
    """
    string_list = [":", ";", ".", ",", "：", "；", "。", "，", "'", "’", " ", "-", "/", "(", ")", "（", "）"]
    for s in string_list:
        text = text.replace(s, "")

    return text


def get_year_list(text: str) -> list[str]:
    """
    Get the year like string using re.
    It will extract all year like strings in format ``YYYY``.

    :param text: Input text
    :type text: str
    :return: Year string list.
    :rtype: list
    """
    pattern = r'\b\d{4}[a-z]?\b'
    return re.findall(pattern, text)


__all__ = ["get_citations_info", "logger", "replace_invalid_char", "get_year_list"]
