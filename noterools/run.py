from .colors import add_cross_ref_style_hook
from .hyperlinks import add_citation_cross_ref_hook
from .word import Word


def run(word_file_path: str, new_file_path: str, isNumbered=False, setColor: int = None, noUnderLine=True,
        set_container_title_italic=True, cross_ref_key_words: list[str] = None):
    """
    Main entry.

    :param word_file_path: Absolute path of your Word document.
    :param new_file_path: Absolute save path of the new Word document. Existed file will be renamed.
    :param isNumbered: If the citation format is numbered.
    :param setColor: Set font color. You can look up the value at `VBA Documentation <https://learn.microsoft.com/en-us/office/vba/api/word.wdcolor>`_.
    :param noUnderLine: If remove the underline of the hyperlink.
    :param set_container_title_italic: If set the container-title and publisher of Chinese paper to Italic.
    :param cross_ref_key_words:
    :return:
    """

    with Word(word_file_path, new_file_path) as word:
        add_citation_cross_ref_hook(word, is_numbered=isNumbered, color=setColor, no_under_line=noUnderLine, set_container_title_italic=set_container_title_italic)
        add_cross_ref_style_hook(word, color=setColor, bold=True, key_word=cross_ref_key_words)
        word.perform()


__all__ = ["run"]
