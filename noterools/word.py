from os.path import basename, dirname, exists
from shutil import move
from typing import Union

from rich.progress import Progress

import pywintypes
from win32com.client import Dispatch, CDispatch, GetObject

from .error import AddBookmarkError, AddHyperlinkError, ContextError
from .utils import logger


class Word:
    """
    Wrapped Word instance.

    """
    def __init__(self, word_file_path: str, save_path: str = None):
        if not exists(word_file_path):
            logger.error(f"File not found: {word_file_path}")
            raise FileNotFoundError(f"File not found: {word_file_path}")

        self.word_file_path = word_file_path
        self.save_path = save_path
        self.word = None
        self.docx = None

        self._context = False

    @property
    def fields(self):
        self._check_context()

        return self.word.Fields

    def __enter__(self):
        try:
            # attach to existed Word progress
            self.word = GetObject(Class="Word.Application")
            self._attached_existed_progress = True

        except pywintypes.com_error:
            self.word = Dispatch("Word.Application")
            self.word.Visible = False
            self._attached_existed_progress = False

        self.docx = self.word.Documents.Open(self.word_file_path)

        self._context = True

        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self._context = False

        self.to_word()
        if not self._attached_existed_progress:
            self.word.Quit()

    def _check_context(self):
        if not self._context or self.word is None or self.docx is None:
            raise ContextError(f"You must enter Noterools context to do following operations")

    def add_bookmark(self, bookmark_id: str, text_range: CDispatch):
        """
        Add a bookmark to Word.

        :param bookmark_id: Unique ID of the bookmark.
        :type bookmark_id: str
        :param text_range: Text range of Word.
        :type text_range:
        :return:
        :rtype:
        """
        self._check_context()

        try:
            self.docx.Bookmarks.Add(Name=bookmark_id, Range=text_range)
        except pywintypes.com_error as error:
            logger.error(f"Cannot add bookmarks: {bookmark_id}. Error is: {error}")
            raise AddBookmarkError(f"Cannot add bookmarks: {bookmark_id}. Error is: {error}")

    def add_hyperlink(self, link_address: Union[str, CDispatch], text_range: CDispatch, tips="", text_to_display="", no_under_line=True):
        """
        Add a hyperlink to Word.

        :param link_address: Internet url or a bookmark, named range, slide number.
        :type link_address:
        :param text_range: Text range of Word.
        :type text_range: CDispatch
        :param tips: Tips to be displayed.
        :type tips: str
        :param text_to_display: If ``text_to_display!=""``, its text will replace that of ``text_range``.
        :type text_to_display: str
        :param no_under_line: If removes the underline style.
        :type no_under_line: bool
        :return:
        :rtype:
        """
        self._check_context()

        kwargs = {
            "Anchor": text_range,
            "Address": "",
            "SubAddress": "",
            "ScreenTip": tips,
            "TextToDisplay": text_to_display,
        }

        if isinstance(link_address, str) and link_address.startswith("http"):
            kwargs["Address"] = link_address

        else:
            kwargs["SubAddress"] = link_address

        try:
            self.docx.Hyperlinks.Add(**kwargs)
        except pywintypes.com_error as error:
            logger.error(f"Cannot add hyperlinks: {link_address}. Error is: {error}")
            raise AddHyperlinkError(f"Cannot add hyperlinks: {link_address}. Error is: {error}")

        if no_under_line:
            text_range.Font.Underline = 0

    def set_save_path(self, save_path: str):
        """
        Set the path the Word will be saved to.

        :param save_path:
        :type save_path:
        :return:
        :rtype:
        """
        self.save_path = save_path

    def to_word(self, new_file_path: str = None):
        """
        Save the Word to a new file.

        :param new_file_path: Absolute path of the new Word document. Existed file will be renamed.
        :type new_file_path: str
        :return:
        :rtype:
        """
        if new_file_path is None:
            if self.save_path is None:
                return
            else:
                new_file_path = self.save_path

        # check the file
        if exists(new_file_path):
            file_basename = basename(new_file_path)
            dir_path = dirname(new_file_path)
            backup_file_name = file_basename.strip(".docx") + "_bak.docx"
            logger.warning(rf"Found existed output file, backup to {dir_path}\{backup_file_name}")

            try:
                move(new_file_path, rf"{dir_path}\{backup_file_name}")
                is_clear = True

            except PermissionError:
                logger.error(f"Can't rename existed file, skip saving")
                is_clear = False

        else:
            is_clear = True

        if is_clear:
            self.docx.SaveAs(new_file_path)


def loop_word_fields(word_obj: Word, code_key_word: str = None, res_key_word: str = None, show_progress=False, progress_text: str = None, ):
    """
    Loop fields in Word.

    :param word_obj: Word object.
    :type word_obj: Word
    :param code_key_word: The key word in field's code which will be used to filter fields.
    :type code_key_word: str
    :param res_key_word: The key word in field's result which will be used to filter fields.
    :type res_key_word: str
    :param show_progress: If shows a friendly progress bar.
    :type show_progress: bool
    :param progress_text: The text of progress bar.
    :type progress_text: str
    :return:
    :rtype:
    """
    if show_progress:

        if progress_text is None:
            progress_text = f"[red]Looping fields in Word...[red]"

        total = len(list(word_obj.fields))

        with Progress() as progress:
            pid = progress.add_task(progress_text, total=total)

            for field in word_obj.fields:
                progress.advance(pid, 1)

                condition1 = False
                condition2 = False

                if code_key_word is None:
                    condition1 = True
                elif code_key_word in field.Code.Text:
                    condition1 = True

                if res_key_word is None:
                    condition2 = True
                elif res_key_word in field.Result.Text:
                    condition2 = True

                if condition1 and condition2:
                    yield field

    else:
        for field in word_obj.fields:

            condition1 = False
            condition2 = False

            if code_key_word is None:
                condition1 = True
            elif code_key_word in field.Code.Text:
                condition1 = True

            if res_key_word is None:
                condition2 = True
            elif res_key_word in field.Result.Text:
                condition2 = True

            if condition1 and condition2:
                yield field


__all__ = ["Word", "loop_word_fields"]
