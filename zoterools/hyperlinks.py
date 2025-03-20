from json import loads
from os.path import basename

import pywintypes
from rich.progress import Progress

from .utils import logger, get_year_list, replace_invalid_char


def add_hyperlinks_to_citations(docx_obj, isNumbered=False, setColor: int = None, noUnderLine=True):
    """
    Add hyperlinks to citations.

    :param docx_obj: Docx object opened by pywin32.
    :param isNumbered: If the citation format is numbered.
    :param setColor: Set font color. You can look up the value at `VBA Documentation <https://learn.microsoft.com/en-us/office/vba/api/word.wdcolor>`_.
    :param noUnderLine: If remove the underline of the hyperlink.
    :return:
    """
    total = len(list(docx_obj.Fields))

    with Progress() as progress:
        pid = progress.add_task(f"[red]Adding hyperlinks..[red]", total=total)

        for field in docx_obj.Fields:

            progress.advance(pid, advance=1)

            if "ADDIN ZOTERO_ITEM" not in field.Code.Text:
                continue

            oRange = field.Result

            if setColor is not None:
                # exclude "(" and ")"
                oRange.MoveStart(Unit=1, Count=1)
                oRange.MoveEnd(Unit=1, Count=-1)
                oRange.Font.Color = setColor
                oRange.MoveStart(Unit=1, Count=-1)
                oRange.MoveEnd(Unit=1, Count=1)

            if isNumbered:
                oRange.Collapse(1)
                oRangeFind = oRange.Find
                oRangeFind.MatchWildcards = True

                # find the number and add hyperlink
                while oRangeFind.Execute("[0-9]{1,}") and oRange.InRange(field.Result):
                    bmtext = f"Ref_{oRange.Text}"
                    docx_obj.Hyperlinks.Add(
                        Anchor=oRange, Address="", SubAddress=bmtext, ScreenTip="",
                        TextToDisplay=""
                        )

                    if noUnderLine:
                        oRange.Font.Underline = 0

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
                                docx_obj.Hyperlinks.Add(
                                    Anchor=oRange, Address="", SubAddress=bmtext,
                                    ScreenTip="", TextToDisplay=""
                                    )
                            except pywintypes.com_error:    # type: ignore
                                logger.error(f"Cannot add hyperlinks: {bmtext}")
                                break

                            if noUnderLine:
                                oRange.Font.Underline = 0

                            is_add_hyperlink = True
                            break

                    if not is_add_hyperlink:
                        text = oRange.Text
                        oRange.MoveStart(Unit=1, Count=-20)
                        oRange.MoveEnd(Unit=1, Count=20)
                        logger.warning(f"Can't set hyperlinks for [{text}] in {oRange.Text}")
                        oRange.MoveStart(Unit=1, Count=20)
                        oRange.MoveEnd(Unit=1, Count=-20)


__all__ = ["add_hyperlinks_to_citations"]
