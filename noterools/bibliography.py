import difflib
import re

from rich.progress import Progress

from .csl import GetCSLJsonHook
from .error import ParamsError
from .hook import HookBase, HOOKTYPE
from .utils import logger
from .word import Word


def _find_page_num_section(text: str) -> list[str]:
    pattern = r"\s\d{1,4}-\d{1,4}[\+\d{1,4}]*[\.,]{1}"
    return re.findall(pattern, text)


def _find_words(text: str, words_list: list[str]):
    pattern = r"\b(?:" + "|".join(re.escape(word) for word in words_list) + r")\b"
    return re.findall(pattern, text)


class BibBookmarkHook(GetCSLJsonHook):
    def __init__(self, is_numbered=False, set_container_title_italic=True, upper_first_char=False, upper_all_words=False, lower_all_words=False, word_dict: list[str] = None):
        super().__init__(name="BibBookmarkHook")
        self.is_numbered = is_numbered
        self.set_container_title_italic = set_container_title_italic
        self._fields_list = []

        if upper_all_words + upper_first_char + lower_all_words >= 2:
            logger.error(f"You must chose only one format rule for article's title.")
            raise ParamsError(f"You must chose only one format rule for article's title.")

        if lower_all_words and word_dict is None:
            logger.error("To prevent proper noun to be lower, you must give your word dictionary contains roper noun.")
            raise ParamsError("To prevent proper noun to be lower, you must give your word dictionary contains roper noun.")

        self.upper_first_char = upper_first_char
        self.upper_all_words = upper_all_words
        self.lower_all_words = lower_all_words
        self.word_dict = word_dict

    def on_iterate(self, word, field):
        if "ADDIN ZOTERO_BIBL" in field.Code.Text:
            self._fields_list.append(field)

        super().on_iterate(word, field)

    def after_iterate(self, word: Word):
        csl_json_dict = self.csl_json_dict
        item_info_tuple = [
            (
                csl_json.get_title(), csl_json.get_container_title(), csl_json.get_author_names(language=csl_json.get_language(defaults="cn"))[0],
                csl_json.get_publisher(), csl_json.get_language(defaults="cn"), item_id
            ) for item_id, csl_json in csl_json_dict.items()
        ]

        for field in self._fields_list:

            # find ZOTERO field.
            if "ADDIN ZOTERO_BIBL" not in field.Code.Text:
                continue

            oRange = field.Result

            # delete existed bookmark
            for oBookMark in oRange.Bookmarks:
                oBookMark.Delete()

            # used for numbered citation
            iCount = 1
            total = len(list(oRange.Paragraphs))

            with Progress() as progress:
                pid = progress.add_task(f"[red]Adding bookmarks..[red]", total=total)

                for oPara in oRange.Paragraphs:
                    progress.advance(pid, advance=1)

                    oRangePara = oPara.Range
                    bmRange = oRangePara

                    if self.is_numbered:
                        bmtext = f"Ref_{iCount}"
                        iCount += 1
                        # these variables need to be checked
                        # let them be "" to avoid UnboundLocalError
                        bib_container_title = ""
                        bib_publisher = ""
                        bib_language = ""

                    else:
                        text = oRangePara.Text
                        bib_title = ""
                        bib_container_title = ""
                        bib_publisher = ""
                        bib_language = ""
                        bib_item_key = ""

                        for index, _tuple in enumerate(item_info_tuple):
                            _title, _container_title, _author, _publisher, _language, _item_id = _tuple
                            if _title in text and _container_title in text and _author in text and f"{_title} " not in text:
                                bib_title = _title
                                bib_container_title = _container_title
                                bib_publisher = _publisher
                                bib_language = _language
                                bib_item_key = _item_id
                                item_info_tuple.pop(index)
                                break

                        if bib_title == "":
                            logger.warning(f"Can't find the corresponding citation of bib: {text}, do you really cite it?")
                            continue

                        bmtext = f"Ref_{bib_item_key}"

                    # set italic for Chinese container title
                    if self.set_container_title_italic and "cn" in bib_language:

                        if bib_container_title != "":
                            split_paragraph = text.split(bib_container_title)
                            pre_paragraph, post_paragraph = split_paragraph[0], split_paragraph[1]
                            bmRange.MoveStart(Unit=1, Count=len(pre_paragraph))
                            bmRange.MoveEnd(Unit=1, Count=-len(post_paragraph))
                            bmRange.Font.Italic = True
                            bmRange.MoveStart(Unit=1, Count=-len(pre_paragraph))
                            bmRange.MoveEnd(Unit=1, Count=len(post_paragraph))

                        if bib_publisher != "":
                            split_paragraph = text.split(bib_publisher)
                            pre_paragraph, post_paragraph = split_paragraph[0], split_paragraph[1]
                            bmRange.MoveStart(Unit=1, Count=len(pre_paragraph))
                            bmRange.MoveEnd(Unit=1, Count=-len(post_paragraph))
                            bmRange.Font.Italic = True
                            bmRange.MoveStart(Unit=1, Count=-len(pre_paragraph))
                            bmRange.MoveEnd(Unit=1, Count=len(post_paragraph))

                    if bib_title != "" and bib_language == "en":
                        if self.upper_all_words:
                            split_paragraph = text.split(bib_title)
                            new_bib_title = bib_title.upper()

                            pre_paragraph, post_paragraph = split_paragraph[0], split_paragraph[1]
                            bmRange.MoveStart(Unit=1, Count=len(pre_paragraph))
                            bmRange.MoveEnd(Unit=1, Count=-len(post_paragraph))
                            bmRange.Text = new_bib_title
                            bmRange.MoveStart(Unit=1, Count=-len(pre_paragraph))
                            bmRange.MoveEnd(Unit=1, Count=len(post_paragraph))

                        elif self.upper_first_char:
                            split_paragraph = text.split(bib_title)
                            new_bib_title = bib_title.split(" ")
                            for index, _word in enumerate(new_bib_title):
                                _word = f"{_word[0].upper()}{_word[1:]}"
                                new_bib_title[index] = _word
                            new_bib_title = " ".join(new_bib_title)

                            pre_paragraph, post_paragraph = split_paragraph[0], split_paragraph[1]
                            bmRange.MoveStart(Unit=1, Count=len(pre_paragraph))
                            bmRange.MoveEnd(Unit=1, Count=-len(post_paragraph))
                            bmRange.Text = new_bib_title
                            bmRange.MoveStart(Unit=1, Count=-len(pre_paragraph))
                            bmRange.MoveEnd(Unit=1, Count=len(post_paragraph))

                        elif self.lower_all_words:
                            res = _find_words(text, self.word_dict)

                            if len(res) == 0:
                                new_bib_title = bib_title.split(" ")
                                for index, _word in enumerate(new_bib_title):
                                    if index == 0:
                                        _word = f"{_word[0].upper()}{_word[1:].lower()}"
                                        new_bib_title[index] = _word
                                    else:
                                        if new_bib_title[index - 1].startswith((":", ".", "?")):
                                            _word = f"{_word[0].upper()}{_word[1:].lower()}"
                                            new_bib_title[index] = _word
                                        else:
                                            new_bib_title[index] = _word.lower()

                                new_bib_title = " ".join(new_bib_title)

                            else:
                                logger.debug(f"Find proper nouns in title: {res}")
                                new_bib_title = bib_title.split(" ")
                                for index, _word in enumerate(new_bib_title):
                                    if index == 0:
                                        _word = f"{_word[0].upper()}{_word[1:].lower()}"
                                        new_bib_title[index] = _word
                                    else:
                                        if new_bib_title[index - 1].endswith((":", ".", "?")):
                                            _word = f"{_word[0].upper()}{_word[1:].lower()}"
                                            new_bib_title[index] = _word
                                        else:
                                            new_bib_title[index] = _word.lower()

                                new_bib_title = " ".join(new_bib_title)

                                for proper_noun in res:
                                    if proper_noun.lower() in new_bib_title:
                                        logger.debug(f"Find proper noun {proper_noun} in title: {new_bib_title}")
                                        new_bib_title = new_bib_title.replace(proper_noun.lower(), proper_noun)

                                    else:
                                        proper_noun_lower = proper_noun.lower()
                                        proper_noun_lower = f"{proper_noun_lower[0].upper()}{proper_noun_lower[1:]}"

                                        if proper_noun_lower in new_bib_title:
                                            logger.debug(f"Find proper noun {proper_noun} in title: {new_bib_title}")
                                            new_bib_title = new_bib_title.replace(proper_noun_lower, proper_noun)

                                        else:
                                            logger.warning(f"Can't find proper noun '{proper_noun}' in title: {new_bib_title}")

                            split_paragraph = text.split(bib_title)
                            logger.debug(f"Update title '{bib_title}' to '{new_bib_title}'")
                            pre_paragraph, post_paragraph = split_paragraph[0], split_paragraph[1]
                            bmRange.MoveStart(Unit=1, Count=len(pre_paragraph))
                            bmRange.MoveEnd(Unit=1, Count=-len(post_paragraph))
                            bmRange.Text = new_bib_title
                            bmRange.MoveStart(Unit=1, Count=-len(pre_paragraph))
                            bmRange.MoveEnd(Unit=1, Count=len(post_paragraph))


                    bmRange.MoveEnd(1, -1)
                    word.add_bookmark(bmtext, bmRange)
                    bmRange.Collapse(0)


class BibUpdateDashSymbolHook(HookBase):
    """
    This hook will replace the dash symbol between page numbers in your bibliography with the right one.

    The default dash symbol between page numbers in the bibliography generated by Zotero is ``-`` (which Unicode is ``002D`` in Times New Roman).
    This hook will replace the dash symbol with ``–``, which Unicode is ``2013``.
    """

    def __init__(self, font_family="Times New Roman", upper_first_char=False, upper_all_words=False, lower_all_words=False):
        super().__init__("BibUpdateDashSymbolHook")
        self.font_family = font_family
        self._fields_list = []

    def on_iterate(self, word, field):
        if "ADDIN ZOTERO_BIBL" in field.Code.Text:
            self._fields_list.append(field)

    def after_iterate(self, word: Word):
        for field in self._fields_list:

            # find ZOTERO field.
            if "ADDIN ZOTERO_BIBL" not in field.Code.Text:
                continue

            field_res_range = field.Result
            total = len(list(field_res_range.Paragraphs))

            with Progress() as progress:
                pid = progress.add_task(f"[red]Update dash symbol..[red]", total=total)

                for index, _bib in enumerate(field_res_range.Paragraphs):
                    progress.advance(pid, advance=1)

                    _bib_range = _bib.Range
                    _bib_text: str = _bib_range.Text

                    # find the page number section
                    res = _find_page_num_section(_bib_text)
                    if len(res) == 0:
                        continue
                    elif len(res) > 1:
                        logger.warning(f"Find multiple page number sections, use the last one: {res}")
                        page_num_section_text = res[-1]
                    else:
                        page_num_section_text = res[0]

                    _bib_text_list = _bib_text.split(page_num_section_text)

                    if len(_bib_text_list) != 2:
                        logger.warning(f"Bibliography should have only one page number section, something is wrong, skip the {index} bibliography...")
                        continue

                    pre_paragraph, post_paragraph = _bib_text_list

                    _bib_range.MoveStart(Unit=1, Count=len(pre_paragraph))
                    _bib_range.MoveEnd(Unit=1, Count=-len(post_paragraph))
                    _bib_range.Text = page_num_section_text.replace("-", "–")
                    _bib_range.Font.Name = self.font_family


def add_update_dash_symbol_hook(word: Word, font_family="Times New Roman"):
    """
    Add hook to replace the dash symbol between page numbers in your bibliography with the right one.

    The default dash symbol between page numbers in the bibliography generated by Zotero is ``-`` (which Unicode is ``002D`` in Times New Roman).
    The hook added will replace the dash symbol with ``–``, which Unicode is ``2013``.

    :param word: ``Word`` object.
    :type word: Word
    :param font_family: The font family you use. Default is "Times New Roman".
    :type font_family: str
    """
    dash_hook = BibUpdateDashSymbolHook(font_family=font_family)
    word.set_hook(dash_hook, hook_type=HOOKTYPE.IN_ITERATE)
    word.set_hook(dash_hook, hook_type=HOOKTYPE.AFTER_ITERATE)


__all__ = ["BibBookmarkHook", "BibUpdateDashSymbolHook", "add_update_dash_symbol_hook"]
