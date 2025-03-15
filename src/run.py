from os.path import exists, basename, dirname
from shutil import move
from time import sleep

from win32com.client import Dispatch

from .bookmark import add_bookmarks_to_bibliography
from .error import ZoteroCitationError
from .hyperlinks import add_hyperlinks_to_citations
from .utils import logger
from .zotero import init_zotero_client


def run(word_file_path: str, new_file_path: str, zotero_id: str, zotero_api_key: str):
    """
    Main entry.

    :param word_file_path: Absolute path of your Word document.
    :param new_file_path: Absolute save path of the new Word document. Existed file will be renamed.
    :param zotero_id: You Zotero ID to connect to Zotero.
    :param zotero_api_key: You Zotero API key to connect to Zotero.
    :return:
    """
    error_flag = False

    # check file
    if exists(new_file_path):
        file_basename = basename(new_file_path)
        dir_path = dirname(new_file_path)
        backup_file_name = file_basename.strip(".docx") + "_bak.docx"
        move(new_file_path, rf"{dir_path}\{backup_file_name}")
        logger.warning(rf"Found existed output file, backup to {dir_path}\{backup_file_name}")

    init_zotero_client(zotero_id, zotero_api_key)

    word = Dispatch("Word.Application")
    word.Visible = False
    docx = word.Documents.Open(word_file_path)

    sleep(1)

    try:
        add_bookmarks_to_bibliography(docx)
        add_hyperlinks_to_citations(docx)
    except (ZoteroCitationError, KeyboardInterrupt):
        error_flag = True

    if not error_flag:
        docx.SaveAs(new_file_path)

    docx.Close(False)
    word.Quit()


__all__ = ["run"]
