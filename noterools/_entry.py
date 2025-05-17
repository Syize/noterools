from .bibliography import add_bib_bookmark_hook
from .citation import add_citation_hyperlink_hook
from .word import Word


def add_citation_cross_ref_hook(word: Word, is_numbered=False, color=16711680, no_under_line=True, set_container_title_italic=True, full_citation_hyperlink=False):
    """
    Register hooks to add hyperlinks from citations to bibliographies.

    :param word: ``noterools.word.Word`` object.
    :type word: Word
    :param is_numbered: If your citation is numbered. Defaults to False.
    :type is_numbered: bool
    :param color: Set font color. Accepts integer decimal value (e.g., 16711680 for blue), 
                 RGB string (e.g., "255, 0, 0" for red), or "word_auto" for automatic color.
                 Defaults to blue (16711680).
                 You can look up the values at `VBA Documentation
                 <https://learn.microsoft.com/en-us/office/vba/api/word.wdcolor>`_.
    :type color: Union[int, str]
    :param no_under_line: If remove the underline of hyperlinks. Defaults to True.
    :type no_under_line: bool
    :param set_container_title_italic: If italicize the container title and publisher name in bibliography. Defaults to True.
    :type set_container_title_italic: bool
    :param full_citation_hyperlink: If True, the entire citation (author and year) will be hyperlinked for the first reference in multiple citations. For subsequent references in the same citation block, only the year will be hyperlinked due to technical limitations. Defaults to False (only year is hyperlinked).
    :type full_citation_hyperlink: bool
    """
    add_citation_hyperlink_hook(word, is_numbered, color, no_under_line, full_citation_hyperlink)
    add_bib_bookmark_hook(word, is_numbered, set_container_title_italic)


__all__ = ["add_citation_cross_ref_hook"]
