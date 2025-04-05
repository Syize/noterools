from .bibliography import add_bib_bookmark_hook
from .citation import add_citation_hyperlink_hook
from .word import Word


def add_citation_cross_ref_hook(word: Word, is_numbered=False, color: int = 16711680, no_under_line=True, set_container_title_italic=True):
    """
    Register hooks to add hyperlinks from citations to bibliographies.

    :param word: ``noterools.word.Word`` object.
    :type word: Word
    :param is_numbered: If your citation is numbered. Defaults to False.
    :type is_numbered: bool
    :param color: Set font color. Defaults to ``blue (16711680)``. You can look up the value at `VBA Documentation
                  <https://learn.microsoft.com/en-us/office/vba/api/word.wdcolor>`_.
    :type color: int
    :param no_under_line: If remove the underline of hyperlinks. Defaults to True.
    :type no_under_line: bool
    :param set_container_title_italic: If italicize the container title and publisher name in bibliography. Defaults to True.
    :type set_container_title_italic: bool
    """
    add_citation_hyperlink_hook(word, is_numbered, color, no_under_line)
    add_bib_bookmark_hook(word, is_numbered, set_container_title_italic)


__all__ = ["add_citation_cross_ref_hook"]
