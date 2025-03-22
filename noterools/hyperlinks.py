from json import loads
from os.path import basename

from .bookmark import BibBookmarkHook, GetCitationInfoHook
from .error import AddHyperlinkError
from .hook import HOOKTYPE, HookBase
from .utils import get_year_list, logger, replace_invalid_char
from .word import Word


class CitationHyperlinkHook(HookBase):
    """
    Hook to add hyperlinks to citations.
    """
    def __init__(self, is_numbered=False, color: int = None, no_under_line=True):
        super().__init__("CitationHyperlinkHook")
        self.is_numbered = is_numbered
        self.color = color
        self.no_under_line = no_under_line

    def on_iterate(self, word_obj: Word, field):
        if "ADDIN ZOTERO_ITEM" not in field.Code.Text:
            return

        # we will change color after add hyperlinks
        color_range = field.Result
        oRange = field.Result

        if self.is_numbered:
            oRange.Collapse(1)
            oRangeFind = oRange.Find
            oRangeFind.MatchWildcards = True

            # find the number and add hyperlink
            while oRangeFind.Execute("[0-9]{1,}") and oRange.InRange(field.Result):
                bmtext = f"Ref_{oRange.Text}"
                word_obj.add_hyperlink(bmtext, oRange, no_under_line=self.no_under_line)
                oRange.Collapse(0)

        else:
            field_value: str = field.Code.Text.strip()
            field_value = field_value.strip("ADDIN ZOTERO_ITEM CSL_CITATION").strip()
            field_value_json = loads(field_value)
            citations_list = field_value_json["citationItems"]

            citation_text = oRange.Text
            citation_text_left = citation_text
            years_list = get_year_list(citation_text)
            citation_text_length = len(citation_text)

            is_first = True
            last_authors_text = ""
            for _year in years_list:

                authors_text = citation_text_left.split(_year)[0]
                if len(replace_invalid_char(authors_text)) < 1:
                    multiple_article_for_one_author = True
                else:
                    last_authors_text = authors_text
                    multiple_article_for_one_author = False

                citation_text_left = citation_text_left[len(authors_text + _year):]

                # move range to the next year string
                if is_first:
                    oRange.MoveStart(Unit=1, Count=len(authors_text))
                    oRange.MoveEnd(Unit=1, Count=-len(citation_text_left))
                    is_first = False

                else:
                    # 5 just works, don't know why.
                    oRange.MoveEnd(Unit=1, Count=len(authors_text) + 5)
                    oRange.MoveStart(Unit=1, Count=len(authors_text) + 4)

                is_add_hyperlink = False
                for _citation in citations_list:
                    citation_year = _citation["itemData"]["issued"]["date-parts"][0][0]

                    if "language" not in _citation["itemData"]:
                        language = "en"
                    else:
                        language = _citation["itemData"]["language"]

                    author_name = _citation["itemData"]["author"][0]
                    if "family" in author_name:
                        if "cn" in language.lower():
                            author_name = author_name["family"] + author_name["given"]
                        else:
                            author_name = author_name["family"]
                    else:
                        author_name = author_name["literal"]

                    if multiple_article_for_one_author:
                        authors_text = last_authors_text

                    _year_without_character = _year[:4]

                    # check the condition
                    res1 = author_name in authors_text and _year_without_character in citation_year
                    res2 = replace_invalid_char(authors_text) == "" and _year_without_character in citation_year
                    res3 = citation_text_length <= 7

                    if res1 or res2 or res3:
                        item_key = basename(_citation["uris"][0])
                        bmtext = f"Ref_{item_key}"

                        try:
                            word_obj.add_hyperlink(bmtext, oRange, no_under_line=self.no_under_line)
                            is_add_hyperlink = True
                            break
                        except AddHyperlinkError:
                            is_add_hyperlink = False

                        break

                if not is_add_hyperlink:
                    text = oRange.Text
                    oRange.MoveStart(Unit=1, Count=-20)
                    oRange.MoveEnd(Unit=1, Count=20)
                    logger.warning(f"Can't set hyperlinks for [{text}] in {oRange.Text}")
                    oRange.MoveStart(Unit=1, Count=20)
                    oRange.MoveEnd(Unit=1, Count=-20)

        if self.color is not None:
            # exclude "(" and ")"
            color_range.MoveStart(Unit=1, Count=1)
            color_range.MoveEnd(Unit=1, Count=-1)
            color_range.Font.Color = self.color
            color_range.MoveStart(Unit=1, Count=-1)
            color_range.MoveEnd(Unit=1, Count=1)


def add_citation_cross_ref_hook(word: Word, is_numbered=False, color: int = None, no_under_line = True, set_container_title_italic=True):
    """
    Add hook to add hyperlinks from citations to bibliographies.

    :param word:
    :type word:
    :param is_numbered:
    :type is_numbered:
    :param color:
    :type color:
    :param no_under_line:
    :type no_under_line:
    :param set_container_title_italic:
    :type set_container_title_italic:
    :return:
    :rtype:
    """
    # with GetCitationInfoHook we can get more info about articles to add proper bookmarks
    get_citations_info_hook = GetCitationInfoHook()
    word.set_hook(get_citations_info_hook)
    word.set_hook(CitationHyperlinkHook(is_numbered, color, no_under_line))
    # add bookmarks after creating hyperlinks is ok
    bib_bookmark_hook = BibBookmarkHook(get_citations_info_hook, is_numbered, set_container_title_italic)
    word.set_hook(bib_bookmark_hook)
    word.set_hook(bib_bookmark_hook, HOOKTYPE.AFTER_ITERATE)


__all__ = ["add_citation_cross_ref_hook", "CitationHyperlinkHook"]
