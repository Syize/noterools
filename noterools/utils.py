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


def parse_color(color_input):
    """
    Parse different color input formats and return the Word VBA Decimal color value.
    
    Accepts:
    - Integer decimal value (e.g., 16711680 for blue)
    - RGB string (e.g., "255, 0, 0" for red)
    - Named color constant (e.g., "word_auto" for automatic color)
    
    :param color_input: Color in various formats
    :type color_input: Union[int, str, None]
    :return: Word VBA Decimal color value
    :rtype: int or None
    """
    # If None, return None (keep default behavior)
    if color_input is None:
        return None
        
    # If already an integer, assume it's already a valid Decimal color value
    if isinstance(color_input, int):
        return color_input
        
    # If string, check for different formats
    if isinstance(color_input, str):
        # Check for named constants
        if color_input.lower() == "word_auto":
            return -16777216  # wdColorAutomatic
            
        # Check for RGB format (e.g., "255, 0, 0")
        try:
            # Split by comma and strip whitespace
            rgb_values = [int(x.strip()) for x in color_input.split(",")]
            
            # If we have 3 values between 0-255, treat as RGB
            if len(rgb_values) == 3 and all(0 <= x <= 255 for x in rgb_values):
                r, g, b = rgb_values
                # Convert to Word Decimal format: B × 2^16 + G × 2^8 + R
                return (b << 16) + (g << 8) + r
        except (ValueError, TypeError):
            pass
            
    # If we reach here, the input format is invalid
    raise ValueError(f"Invalid color format: {color_input}. Use an integer Decimal value, RGB string (e.g., '255, 0, 0'), or 'word_auto'.")


__all__ = ["logger", "replace_invalid_char", "get_year_list", "find_urls", "parse_color"]
