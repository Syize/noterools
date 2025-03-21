from .bookmark import add_bookmarks_to_bibliography
from .hyperlinks import add_hyperlinks_to_citations
from .word import Word
from .zotero import init_zotero_client


def run(word_file_path: str, new_file_path: str, zotero_id: str, zotero_api_key: str, isNumbered=False, setColor: int = None, noUnderLine=True,
        set_container_title_italic=True):
    """
    Main entry.

    :param word_file_path: Absolute path of your Word document.
    :param new_file_path: Absolute save path of the new Word document. Existed file will be renamed.
    :param zotero_id: You Zotero ID to connect to Zotero.
    :param zotero_api_key: You Zotero API key to connect to Zotero.
    :param isNumbered: If the citation format is numbered.
    :param setColor: Set font color. You can look up the value at `VBA Documentation <https://learn.microsoft.com/en-us/office/vba/api/word.wdcolor>`_.
    :param noUnderLine: If remove the underline of the hyperlink.
    :param set_container_title_italic: If set the container-title and publisher of Chinese paper to Italic.
    :return:
    """
    init_zotero_client(zotero_id, zotero_api_key)

    with Word(word_file_path, new_file_path) as word:
        add_bookmarks_to_bibliography(word, isNumbered=isNumbered, set_container_title_italic=set_container_title_italic)
        add_hyperlinks_to_citations(word, isNumbered=isNumbered, setColor=setColor, noUnderLine=noUnderLine)


__all__ = ["run"]
