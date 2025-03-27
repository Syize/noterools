from json import loads
from os.path import basename

from rich.progress import Progress

from .hook import HookBase
from .utils import logger
from .word import Word


class GetCitationInfoHook(HookBase):
    """
    Get article info from citations.
    """
    def __init__(self):
        super().__init__("GetCitationInfoHook")
        self.titles_item_keys = {}

    def get_citations_info(self):
        """

        :return:
        :rtype:
        """
        return self.titles_item_keys

    def on_iterate(self, word, field):
        if "ADDIN ZOTERO_ITEM" not in field.Code.Text:
            return

        # convert string to JSON string.
        field_value: str = field.Code.Text.strip()
        field_value = field_value.strip("ADDIN ZOTERO_ITEM CSL_CITATION").strip()
        field_value_json = loads(field_value)
        citations_list = field_value_json["citationItems"]

        for _citation in citations_list:
            # pprint(_citation)
            # raise ZoteroCitationError
            item_key = basename(_citation["uris"][0])
            title = _citation["itemData"]["title"]

            if "container-title" in _citation["itemData"]:
                container_title = _citation["itemData"]["container-title"]
            else:
                container_title = ""

            if "publisher" in _citation["itemData"]:
                publisher = _citation["itemData"]["publisher"]
            else:
                publisher = ""

            if "language" not in _citation["itemData"]:
                language = "cn"
            else:
                language = _citation["itemData"]["language"]
                if "en" not in language.lower():
                    language = "cn"
                else:
                    language = "en"

            author = _citation["itemData"]["author"][0]
            if "family" in author:
                author = author["family"]
            else:
                author = author["literal"]

            if title not in self.titles_item_keys:
                self.titles_item_keys[title] = {
                    "item_key": item_key,
                    "container_title": container_title,
                    "author": author,
                    "publisher": publisher,
                    "language": language,
                }


class BibBookmarkHook(HookBase):
    def __init__(self, citation_info_hook: GetCitationInfoHook, is_numbered=False, set_container_title_italic=True):
        super().__init__("BibBookmarkHook")
        self.citation_info_hook = citation_info_hook
        self.is_numbered = is_numbered
        self.set_container_title_italic = set_container_title_italic
        self._fields_list = []

    def on_iterate(self, word, field):
        if "ADDIN ZOTERO_BIBL" in field.Code.Text:
            self._fields_list.append(field)

    def after_iterate(self, word: Word):
        title_item_key_dict = self.citation_info_hook.get_citations_info()
        title_container_title_tuple = [
            (
                title, title_item_key_dict[title]["container_title"], title_item_key_dict[title]["author"], title_item_key_dict[title]["publisher"],
                title_item_key_dict[title]["language"],
            ) for title in title_item_key_dict
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

                        for index, _tuple in enumerate(title_container_title_tuple):
                            _title, _container_title, _author, _publisher, _language = _tuple
                            if _title in text and _container_title in text and _author in text and f"{_title} " not in text:
                                bib_title = _title
                                bib_container_title = _container_title
                                bib_publisher = _publisher
                                bib_language = _language
                                title_container_title_tuple.pop(index)
                                break

                        if bib_title == "":
                            logger.warning(f"Can't find the corresponding citation of bib: {text}, do you really cite it?")
                            continue

                        bib_item_key = title_item_key_dict.pop(bib_title)["item_key"]
                        bmtext = f"Ref_{bib_item_key}"

                    # set italic for Chinese container title
                    if self.set_container_title_italic and bib_language == "cn":

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

                    bmRange.MoveEnd(1, -1)
                    word.add_bookmark(bmtext, bmRange)
                    bmRange.Collapse(0)


__all__ = ["GetCitationInfoHook", "BibBookmarkHook"]
