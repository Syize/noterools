from os.path import exists

from win32com.client import Dispatch, CDispatch

from .utils import logger


def open_word(word_file_path: str) -> CDispatch:
    """

    :param word_file_path: Absolute path of your Word document.
    :type word_file_path: str
    :return: Docx object opened by pywin32.
    :rtype: CDispatch
    """
    if not exists(word_file_path):
        logger.error(f"File not found: {word_file_path}")
        raise FileNotFoundError(f"File not found: {word_file_path}")

    word = Dispatch("Word.Application")
    word.Visible = False
    docx = word.Documents.Open(word_file_path)

    return docx


__all__ = ["open_word"]
