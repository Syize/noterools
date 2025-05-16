# -*- coding: utf-8 -*-
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


def find_urls(text: str) -> list[tuple[int, int, str]]:
    """
    Find URLs in text and return their positions and values.
    
    :param text: The text to search
    :type text: str
    :return: List of tuples (start_pos, end_pos, url)
    :rtype: list[tuple[int, int, str]]
    """
    # Pattern to match common URL formats, excluding trailing punctuation
    url_pattern = r'(https?://(?:www\.)?[-a-zA-Z0-9@:%._\+~#=]{1,256}\.[a-zA-Z0-9()]{1,6}\b(?:[-a-zA-Z0-9()@:%_\+.~#?&//=]*))'
    
    # Pattern to match DOIs, excluding trailing punctuation
    doi_pattern = r'(doi\.org/[0-9a-zA-Z./\-_]+)'
    
    # Combine patterns
    combined_pattern = f"{url_pattern}|{doi_pattern}"
    
    urls = []
    for match in re.finditer(combined_pattern, text):
        start, end = match.span()
        url = match.group(0)
        
        # Remove trailing punctuation
        while url and url[-1] in '.,:;)]}"\'':
            url = url[:-1]
            end -= 1
        
        if url:  # Only add if URL is not empty after processing
            urls.append((start, end, url))
    
    return urls


__all__ = ["logger", "replace_invalid_char", "get_year_list", "find_urls"]
