import logging
import re

from rich.logging import RichHandler

logger = logging.getLogger("noterools")
formatter = logging.Formatter("%(name)s :: %(message)s", datefmt="%m-%d %H:%M:%S")
handler = RichHandler()
handler.setFormatter(formatter)
logger.addHandler(handler)
logger.setLevel(logging.INFO)


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


__all__ = ["logger", "replace_invalid_char", "get_year_list"]
