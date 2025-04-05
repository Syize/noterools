import re
from typing import Optional

from rich.progress import Progress

from .csl import GetCSLJsonHook, add_get_csl_json_hook
from .error import ParamsError, HookNotRegisteredError
from .hook import ExtensionHookBase, HOOKTYPE, HookBase
from .utils import logger
from .word import Word


def _find_page_num_section(text: str) -> list[str]:
    pattern = r"\s\d{1,4}-\d{1,4}[\+\d{1,4}]*[\.,]{1}"
    return re.findall(pattern, text)


def _find_words(text: str, words_list: list[str]):
    pattern = r"\b(?:" + "|".join(re.escape(word) for word in words_list) + r")\b"
    return re.findall(pattern, text)


def _check_get_csl_json_hook(hook_name: str, hook: GetCSLJsonHook):
    if not hook.is_registered():
        logger.error(f"Hook '{hook_name}' requires `GetCSLJsonHook` to be registered. Use function `add_get_csl_json_hook` to register it.")
        raise HookNotRegisteredError(f"Hook '{hook_name}' requires `GetCSLJsonHook` to be registered. Use function `add_get_csl_json_hook` to register it.")


def italicize_container_publisher(paragraph_range, container_title="", publisher=""):
    """
    Italicize the container title and publisher name in bibliography text.

    :param paragraph_range: Range object of the bibliography.
    :type paragraph_range:
    :param container_title: Container title.
    :type container_title: str
    :param publisher: Publisher name.
    :type publisher: str
    :return:
    :rtype:
    """
    bib_text = paragraph_range.Text
    if container_title != "":
        split_paragraph = bib_text.split(container_title)
        pre_paragraph, post_paragraph = split_paragraph[0], split_paragraph[1]
        paragraph_range.MoveStart(Unit=1, Count=len(pre_paragraph))
        paragraph_range.MoveEnd(Unit=1, Count=-len(post_paragraph))
        paragraph_range.Font.Italic = True
        paragraph_range.MoveStart(Unit=1, Count=-len(pre_paragraph))
        paragraph_range.MoveEnd(Unit=1, Count=len(post_paragraph))

    if publisher != "":
        split_paragraph = bib_text.split(publisher)
        pre_paragraph, post_paragraph = split_paragraph[0], split_paragraph[1]
        paragraph_range.MoveStart(Unit=1, Count=len(pre_paragraph))
        paragraph_range.MoveEnd(Unit=1, Count=-len(post_paragraph))
        paragraph_range.Font.Italic = True
        paragraph_range.MoveStart(Unit=1, Count=-len(pre_paragraph))
        paragraph_range.MoveEnd(Unit=1, Count=len(post_paragraph))


class BibLoopHook(HookBase):
    """
    This hook loops each paragraph in bibliography, and calls extension hook in each iteration.
    Extension hook is a type of hook to process paragraphs in bibliography.
    See ``noterools.hook.ExtensionHookBase``.
    """

    def __init__(self):
        if self._hook_initialized:
            return

        super().__init__(name="BibLoopHook")
        self._hook_dict: dict[str, ExtensionHookBase] = {}
        self._fields_list = []

    def set_hook(self, hook: ExtensionHookBase):
        """
        Set extension hook.
        Only accept HOOKTYPE.IN_ITERATE type.

        :param hook: ExtensionHookBase type hook.
        :type hook: ExtensionHookBase
        :return:
        :rtype:
        """
        if hook.is_registered():
            return

        if hook.name in self._hook_dict:
            logger.warning(f"Hook {hook.name} won't be added because a hook with same name exists.")
            return

        self._hook_dict.update({hook.name: hook})
        hook.finish_register()

    def update_hook(self, hook: ExtensionHookBase):
        """
        Update specific child hook.
        Only accept HOOKTYPE.IN_ITERATE type.

        :param hook: ExtensionHookBase type hook.
        :type hook: ExtensionHookBase
        :return:
        :rtype:
        """
        if hook.is_registered():
            return

        if hook.name in self._hook_dict:
            logger.warning(f"Hook {hook.name} doesn't exist.")
            return

        self._hook_dict.update({hook.name: hook})
        hook.finish_register()

    def remove_hook(self, name: str):
        """
        Remove a hook.

        :param name: Hook's name.
        :type name: str
        :return:
        :rtype:
        """
        if name in self._hook_dict:
            self._hook_dict.pop(name)

    def on_iterate(self, word, field):
        if "ADDIN ZOTERO_BIBL" in field.Code.Text:
            self._fields_list.append(field)

    def after_iterate(self, word: Word):
        if len(self._fields_list) == 0:
            logger.warning(f"Zotero's bibliography not found in your file. Generate bibliography in your file, and DO NOT unlink it.")
            return

        elif len(self._fields_list) > 1:
            logger.warning(f"Found multiple bibliography in your file, this is weird.")

        for field in self._fields_list:

            # find ZOTERO field.
            if "ADDIN ZOTERO_BIBL" not in field.Code.Text:
                logger.debug(f"Exclude field that is not Zotero bibliography: {field.Code.Text}")
                continue

            field_res_range = field.Result

            # used for numbered citation
            total = len(list(field_res_range.Paragraphs))

            with Progress() as progress:
                pid = progress.add_task(f"[red]Processing your bibliography..[red]", total=total)

                for _paragraph in field_res_range.Paragraphs:
                    progress.advance(pid, advance=1)

                    for _hook_name in self._hook_dict:
                        logger.debug(f"Call hook: {_hook_name}")
                        self._hook_dict[_hook_name].on_iterate(word, _paragraph.Range)


def add_bib_loop_hook(word: Word) -> BibLoopHook:
    """
    Register ``BibLoopHook``.

    :param word: ``noterools.word.Word`` object.
    :type word: Word
    :return: ``BibLoopHook`` instance.
    :rtype: BibLoopHook
    """
    bib_loop_hook = BibLoopHook()
    word.set_hook(bib_loop_hook, HOOKTYPE.IN_ITERATE)
    word.set_hook(bib_loop_hook, HOOKTYPE.AFTER_ITERATE)

    return bib_loop_hook


class BibBookmarkHook(ExtensionHookBase):
    def __init__(self, is_numbered=False, set_container_title_italic=True):
        """
        This extension hook can add bookmark to your bibliography so you can create hyperlinks from citations to bibliography.

        :param is_numbered: If your citation is numbered. Defaults to False.
        :type is_numbered: bool
        :param set_container_title_italic: If italicize the container title and publisher name in bibliography. Defaults to True.
        :type set_container_title_italic: bool
        """
        super().__init__(name="BibBookmarkHook")
        self.is_numbered = is_numbered
        self.set_container_title_italic = set_container_title_italic
        self._fields_list = []
        self._get_cls_json_hook = GetCSLJsonHook()

        # we need ``GetCSLJsonHook`` to be registered.
        _check_get_csl_json_hook(self.name, self._get_cls_json_hook)

        # used to generate bookmark for numbered citation
        self._number_count = 1

        # used to match the citation with bibliography.
        self._item_info_list: Optional[list[tuple[str, str, str, str, str, str]]] = None

    def _get_bookmark_id_and_item_info(self, bib_text: str) -> tuple[str, tuple[str, str, str, str]]:
        """
        Get bookmark id and information about the article.

        :param bib_text: Text of the bibliography.
        :type bib_text: str
        :return: (bookmark_id, (title, container_title, publisher, language))
        :rtype: tuple
        """
        if self._item_info_list is None:
            csl_json_dict = self._get_cls_json_hook.get_csl_jsons()
            # [("title", "container title", "first author name", "publisher", "language", "item id"), ...]
            self._item_info_list = [
                (
                    csl_json.get_title(), csl_json.get_container_title(), csl_json.get_author_names(language=csl_json.get_language(defaults="cn"))[0],
                    csl_json.get_publisher(), csl_json.get_language(defaults="cn"), item_id
                ) for item_id, csl_json in csl_json_dict.items()
            ]

        bib_title = ""
        bib_container_title = ""
        bib_publisher = ""
        bib_language = ""
        bib_item_id = ""

        for index, _tuple in enumerate(self._item_info_list):
            _title, _container_title, _author, _publisher, _language, _item_id = _tuple

            # we have to check following things to make sure this is the article we find for bibliography
            # 1. bib text contains article's title.
            # 2. bib text contains article's container title (container title will be `""` if your Zotero doesn't have information about it).
            # 3. bib text contains the first author's name.
            # 4. article's title must match the title in bib text perfectly.
            if _title in bib_text and _container_title in bib_text and _author in bib_text and f"{_title} " not in bib_text:
                bib_title = _title
                bib_container_title = _container_title
                bib_publisher = _publisher
                bib_language = _language
                bib_item_id = _item_id
                self._item_info_list.pop(index)
                break

        article_info = (bib_title, bib_container_title, bib_publisher, bib_language)

        if self.is_numbered:
            bookmark_id = f"Ref_{self._number_count}"
            self._number_count += 1

        else:
            # item id is unique in Zotero
            bookmark_id = f"Ref_{bib_item_id}"

        return bookmark_id, article_info

    def on_iterate(self, word: Word, word_range):
        bib_text = word_range.Text
        bookmark_id, (bib_title, bib_container_title, bib_publisher, bib_language) = self._get_bookmark_id_and_item_info(bib_text)

        if not self.is_numbered and bib_title == "":
            logger.warning(f"Can't find the corresponding citation of bib: {bib_text}, do you really cite it?")
            return

        # set italic for Chinese container title
        if self.set_container_title_italic and "cn" in bib_language:
            italicize_container_publisher(word_range, bib_container_title, bib_publisher)

        word_range.MoveEnd(1, -1)
        word.add_bookmark(bookmark_id, word_range)


def add_bib_bookmark_hook(word: Word, is_numbered=False, set_container_title_italic=True) -> BibBookmarkHook:
    """
    Register ``BibBookmarkHook``.

    :param word: ``noterools.word.Word`` object.
    :type word: Word
    :param is_numbered: If your citation is numbered. Defaults to False.
    :type is_numbered: bool
    :param set_container_title_italic: If italicize the container title and publisher name in bibliography. Defaults to True.
    :type set_container_title_italic: bool
    :return: ``BibBookmarkHook`` instance.
    :rtype: BibBookmarkHook
    """
    add_get_csl_json_hook(word)
    bib_bookmark_hook = BibBookmarkHook(is_numbered, set_container_title_italic)
    bib_loop_hook = add_bib_loop_hook(word)
    bib_loop_hook.set_hook(bib_bookmark_hook)

    return bib_bookmark_hook


class BibUpdateDashSymbolHook(ExtensionHookBase):
    def __init__(self, font_family="Times New Roman"):
        """
        This hook will replace the dash symbol between page numbers in your bibliography with the right one.

        The default dash symbol between page numbers in the bibliography generated by Zotero is ``-`` (which Unicode is ``002D`` in Times New Roman).
        This hook will replace the dash symbol with ``–``, which Unicode is ``2013``.

        :param font_family: The font you use. Defaults to ``Times New Roman``.
        :type font_family: 
        """
        super().__init__("BibUpdateDashSymbolHook")
        self.font_family = font_family
        self._fields_list = []

    def on_iterate(self, word: Word, word_range):
        _bib_text: str = word_range.Text

        # find the page number section
        res = _find_page_num_section(_bib_text)
        if len(res) == 0:
            return
        elif len(res) > 1:
            logger.warning(f"Find multiple page number sections, use the last one: {res}")
            page_num_section_text = res[-1]
        else:
            page_num_section_text = res[0]

        _bib_text_list = _bib_text.split(page_num_section_text)

        if len(_bib_text_list) != 2:
            logger.warning(f"Bibliography should have only one page number section, something is wrong, skip the paragraph: {_bib_text}")
            return

        pre_paragraph, post_paragraph = _bib_text_list

        word_range.MoveStart(Unit=1, Count=len(pre_paragraph))
        word_range.MoveEnd(Unit=1, Count=-len(post_paragraph))
        word_range.Text = page_num_section_text.replace("-", "–")
        word_range.Font.Name = self.font_family


def add_update_dash_symbol_hook(word: Word, font_family="Times New Roman") -> BibUpdateDashSymbolHook:
    """
    Register ``BibUpdateDashSymbolHook``.

    :param word: ``noterools.word.Word`` object.
    :type word: Word
    :param font_family: The font family you use. Default is "Times New Roman".
    :type font_family: str
    :return: ``BibUpdateDashSymbolHook`` instance.
    :rtype: BibUpdateDashSymbolHook
    """
    bib_update_dash_symbol_hook = BibUpdateDashSymbolHook(font_family)
    bib_loop_hook = add_bib_loop_hook(word)
    bib_loop_hook.set_hook(bib_update_dash_symbol_hook)

    return bib_update_dash_symbol_hook


class BibFormatTitleHook(ExtensionHookBase):
    def __init__(self, upper_first_char=False, upper_all_words=False, lower_all_words=False, word_list: list[str] = None):
        """
        This hook will format the article title with three rules: upper_all, upper_first_character, and lower_all.

        :param upper_first_char: Upper the first character of each word. Defaults to False.
        :type upper_first_char: bool
        :param upper_all_words: Upper all words. Defaults to False.
        :type upper_all_words: bool
        :param lower_all_words: Lower characters except the first in all words. Words or phrases in ``word_dict`` will not be changed. Defaults to False.
        :type lower_all_words: bool
        :param word_list: A list contains words or phrases which format will not be changed.
        :type word_list: list
        """
        super().__init__("BibFormatTitleHook")
        self._get_cls_json_hook = GetCSLJsonHook()

        # we need ``GetCSLJsonHook`` to be registered.
        _check_get_csl_json_hook(self.name, self._get_cls_json_hook)

        # used to match the citation with bibliography.
        self._item_info_list: Optional[list[tuple[str, str, str, str, str, str]]] = None
        
        if upper_all_words + upper_first_char + lower_all_words >= 2:
            logger.error(f"You must chose only one format rule for article's title.")
            raise ParamsError(f"You must chose only one format rule for article's title.")

        if lower_all_words and word_list is None:
            logger.error("To prevent proper noun to be lower, you must give your word dictionary contains roper noun.")
            raise ParamsError("To prevent proper noun to be lower, you must give your word dictionary contains roper noun.")

        self.upper_first_char = upper_first_char
        self.upper_all_words = upper_all_words
        self.lower_all_words = lower_all_words

        # remove empty strings
        word_list = [x for x in word_list if x != ""]
        self.word_list = word_list

        logger.warning("Change the capitalization of article title is an experimental feature and may make mistakes.")
        logger.warning("Use it carefully.")
        
    def _get_title(self, bib_text: str) -> tuple[str, str]:
        """
        Get the title and language of a bibliography.

        :param bib_text: Text of the bibliography.
        :type bib_text: str
        :return: (title, language)
        :rtype: tuple
        """
        if self._item_info_list is None:
            csl_json_dict = self._get_cls_json_hook.get_csl_jsons()
            # [("title", "container title", "first author name", "publisher", "language", "item id"), ...]
            self._item_info_list = [
                (
                    csl_json.get_title(), csl_json.get_container_title(), csl_json.get_author_names(language=csl_json.get_language(defaults="cn"))[0],
                    csl_json.get_publisher(), csl_json.get_language(defaults="cn"), item_id
                ) for item_id, csl_json in csl_json_dict.items()
            ]

        bib_title = ""
        bib_language = ""

        for index, _tuple in enumerate(self._item_info_list):
            _title, _container_title, _author, _publisher, _language, _item_id = _tuple

            # we have to check following things to make sure this is the article we find for bibliography
            # 1. bib text contains article's title.
            # 2. bib text contains article's container title (container title will be `""` if your Zotero doesn't have information about it).
            # 3. bib text contains the first author's name.
            # 4. article's title must match the title in bib text perfectly.
            if _title in bib_text and _container_title in bib_text and _author in bib_text and f"{_title} " not in bib_text:
                bib_title = _title
                bib_language = _language
                self._item_info_list.pop(index)
                break

        return bib_title, bib_language

    def on_iterate(self, word: Word, word_range):
        bib_text = word_range.Text
        bib_title, bib_language = self._get_title(bib_text)

        if bib_title != "" and bib_language == "en":
            if self.upper_all_words:
                split_paragraph = bib_text.split(bib_title)
                new_bib_title = bib_title.upper()

                pre_paragraph, post_paragraph = split_paragraph[0], split_paragraph[1]
                word_range.MoveStart(Unit=1, Count=len(pre_paragraph))
                word_range.MoveEnd(Unit=1, Count=-len(post_paragraph))
                word_range.Text = new_bib_title
                word_range.MoveStart(Unit=1, Count=-len(pre_paragraph))
                word_range.MoveEnd(Unit=1, Count=len(post_paragraph))

            elif self.upper_first_char:
                split_paragraph = bib_text.split(bib_title)
                new_bib_title = bib_title.split(" ")
                for index, _word in enumerate(new_bib_title):
                    _word = f"{_word[0].upper()}{_word[1:]}"
                    new_bib_title[index] = _word
                new_bib_title = " ".join(new_bib_title)

                pre_paragraph, post_paragraph = split_paragraph[0], split_paragraph[1]
                word_range.MoveStart(Unit=1, Count=len(pre_paragraph))
                word_range.MoveEnd(Unit=1, Count=-len(post_paragraph))
                word_range.Text = new_bib_title
                word_range.MoveStart(Unit=1, Count=-len(pre_paragraph))
                word_range.MoveEnd(Unit=1, Count=len(post_paragraph))

            elif self.lower_all_words:
                res = _find_words(bib_text, self.word_list)
                res = [x for x in res if x != ""]

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

                split_paragraph = bib_text.split(bib_title)
                logger.debug(f"Update title '{bib_title}' to '{new_bib_title}'")
                pre_paragraph, post_paragraph = split_paragraph[0], split_paragraph[1]
                word_range.MoveStart(Unit=1, Count=len(pre_paragraph))
                word_range.MoveEnd(Unit=1, Count=-len(post_paragraph))
                word_range.Text = new_bib_title
                word_range.MoveStart(Unit=1, Count=-len(pre_paragraph))
                word_range.MoveEnd(Unit=1, Count=len(post_paragraph))


def add_format_title_hook(word: Word, upper_first_char=False, upper_all_words=False, lower_all_words=False, word_list: list[str] = None) -> BibFormatTitleHook:
    """
    Register ``BibFormatTitleHook``.

    :param word: ``noterools.word.Word`` object.
    :type word: Word
    :param upper_first_char: Upper the first character of each word. Defaults to False.
    :type upper_first_char: bool
    :param upper_all_words: Upper all words. Defaults to False.
    :type upper_all_words: bool
    :param lower_all_words: Lower characters except the first in all words. Words or phrases in ``word_dict`` will not be changed. Defaults to False.
    :type lower_all_words: bool
    :param word_list: A list contains words or phrases which format will not be changed.
    :type word_list: list
    :return: ``BibFormatTitleHook`` instance.
    :rtype: BibFormatTitleHook
    """
    add_get_csl_json_hook(word)
    bib_format_title_hook = BibFormatTitleHook(upper_first_char, upper_all_words, lower_all_words, word_list)
    bib_loop_hook = add_bib_loop_hook(word)
    bib_loop_hook.set_hook(bib_format_title_hook)

    return bib_format_title_hook


__all__ = ["BibLoopHook", "BibBookmarkHook", "BibUpdateDashSymbolHook", "BibFormatTitleHook", "add_bib_loop_hook", "add_bib_bookmark_hook", "add_update_dash_symbol_hook",
           "add_format_title_hook"]
