import re
from json import loads

from rich.progress import Progress

import pywintypes
from win32com.client import Dispatch


class ZoteroCitationError(Exception):
    """
    A custom exception class
    """
    pass


def replace_invalid_char(text: str) -> str:
    """
    Replace invalid characters with "" because bookmarks in Word mustn't contain these characters.

    :param text: Input text.
    :type text: str
    :return: Text in which all invalid characters have been replaced.
    :rtype: str
    """
    string_list = [":", ";", ".", ",", "：", "；", "。", "，", "'", "’", " ", "-"]
    for s in string_list:
        text = text.replace(s, "")

    return text


def get_year(text: str, extra_pattern: str = "") -> str:
    """
    Get the year like string using re.
    It will match the first year like string in format ``YYYYx``, where ``x`` is a single character in ``a – e``

    :param text: Input text
    :type text: str
    :param extra_pattern: Extra patterns.
    :type extra_pattern: str
    :return: Year string.
    :rtype: str
    """
    pattern = r'\b(19\d{2}|20\d{2})[abcde]?\b' + extra_pattern  # 匹配 1900-2099 之间的四位年份
    res = re.findall(pattern, text)
    if len(res) < 1:
        res_year = ""
    else:
        res_year = res[0]
    return res_year


def get_authors_string(text: str, year_str: str) -> str:
    """
    Get all authors before ``year_str`` and generate a string using all authors' name.

    :param text: Input text.
    :type text: str
    :param year_str: Year of the paper.
    :type year_str: str
    :return: Generated string contains all authors' name.
    :rtype: str
    """
    text = text.split(year_str)[0].strip()
    text = text.replace(",", "_").replace(" ", "").strip("_").strip(".")
    return text


def create_bookmarks_for_zotero_literature(docx_obj, isNumbered=False):
    """
    Parse Zotero bibliographies, add bookmark for each reference.

    :param docx_obj: Docx object opened by pywin32.
    :type docx_obj:
    :param isNumbered: If the citation format is numbered.
    :type isNumbered: bool
    :return:
    :rtype:
    """
    total = len(list(docx_obj.Fields))

    with Progress() as progress:
        pid = progress.add_task(f"[red]Adding bookmarks..[red]", total=total)

        # loop fields in docx
        for field in docx_obj.Fields:

            progress.advance(pid, advance=1)

            # find ZOTERO field.
            if "ADDIN ZOTERO_BIBL" in field.Code.Text:
                oRange = field.Result

                # delete existed bookmark
                for oBookMark in oRange.Bookmarks:
                    oBookMark.Delete()

                # used for numbered citation
                iCount = 1

                for oPara in oRange.Paragraphs:
                    oRangePara = oPara.Range
                    bmRange = oRangePara

                    if isNumbered:
                        bmtext = f"Ref_{iCount}"
                        iCount += 1

                    else:
                        oYear = get_year(oRangePara.Text, r"[;；。，：:.,\)]")
                        if oYear == "":
                            print("GetYear error")
                            print(oRangePara.Text)
                            raise ZoteroCitationError
                        oAuthors = get_authors_string(oRangePara.Text, oYear)
                        # generate unique bookmark based on authors and the year
                        bmtext = "Ref_" + oAuthors + "_" + oYear
                        bmtext = replace_invalid_char(bmtext)

                    bmRange.MoveEnd(1, -1)
                    try:
                        docx_obj.Bookmarks.Add(Name=bmtext, Range=bmRange)
                    except pywintypes.com_error:
                        print(f"Cannot add bookmarks: {bmtext}")
                        raise ZoteroCitationError
                    bmRange.Collapse(0)


def _generate_bookmark_id(authors_list: list[dict[str, str]], etal_text: str, citation_year: str, etal_number: int = None, is_cn_language=False) -> str:
    """
    Generate the target bookmark id based on authors' name, papers' date and etal text.

    :param authors_list: A list contains name information about all authors.
    :type authors_list: list
    :param etal_text: Corresponding text of "et al."
    :type etal_text: str
    :param citation_year: Publish year of the paper.
    :type citation_year: str
    :param etal_number: The number of authors when using "et al." to represent the rest of authors.
    :type etal_number: int | None
    :param is_cn_language: If the language of paper is Chinese.
    :type is_cn_language: bool
    :return: Generated bookmark id.
    :rtype: str
    """
    if etal_number is not None and len(authors_list) > etal_number:
        use_etal = True
        authors_list = authors_list[:etal_number]
    else:
        use_etal = False

    _text = ""
    for _author in authors_list:
        if "family" not in _author:
            _text += _author["literal"].replace(" ", "")
        else:
            if is_cn_language:
                _text += _author["family"] + _author["given"].replace(" ", "") + "_"
            else:
                _author_given = _author["given"]
                if "-" in _author_given:
                    _author_given_list = _author_given.split("-")
                    _author_given_list = [x[0].upper() for x in _author_given_list]
                    _author_given = "".join(_author_given_list)
                else:
                    _author_given = _author["given"][0].upper()

                _text += _author["family"] + _author_given + "_"

    _text = _text.strip("_")

    if use_etal:
        _text += f"_{etal_text}"

    _text = replace_invalid_char(_text)
    bmtext = f"Ref_{_text}_{citation_year}"

    return bmtext


def create_hyperlinks_to_literature_bookmarks(docx_obj, isNumbered=False, setColor: int = None, etal_number: int = None):
    """
    Add hyperlinks to corresponding bookmarks.

    :param docx_obj: Docx object opened by pywin32.
    :type docx_obj:
    :param isNumbered: If the citation format is numbered.
    :type isNumbered: bool
    :param setColor: The id of the color you want to use to change the citation.
    :type setColor: int
    :param etal_number: The number of authors when using "et al." to represent the rest of authors.
    :type etal_number: int | None
    :return:
    :rtype:
    """
    total = len(list(docx_obj.Fields))

    with Progress() as progress:
        pid = progress.add_task(f"[red]Adding hyperlinks..[red]", total=total)

        for field in docx_obj.Fields:

            progress.advance(pid, advance=1)

            if "ADDIN ZOTERO_ITEM" in field.Code.Text:
                oRange = field.Result

                if setColor is not None:
                    # exclude "(" and ")"
                    oRange.MoveStart(Unit=1, Count=1)
                    oRange.MoveEnd(Unit=1, Count=-1)
                    oRange.Font.Color = setColor
                    oRange.MoveStart(Unit=1, Count=-1)
                    oRange.MoveEnd(Unit=1, Count=1)

                oRangeFind = oRange.Find
                oRangeFind.MatchWildcards = True

                if isNumbered:
                    oRange.Collapse(1)

                    # find the number and add hyperlink
                    while oRangeFind.Execute("[0-9]{1,}") and oRange.InRange(field.Result):
                        bmtext = f"Ref_{oRange.Text}"
                        docx_obj.Hyperlinks.Add(Anchor=oRange, Address="", SubAddress=bmtext, ScreenTip="",
                                                TextToDisplay="")
                        oRange.Collapse(0)

                else:

                    # we need to generate the same name so we can link to the corresponding bookmark.
                    # load information from the field code.
                    field_value: str = field.Code.Text.strip()
                    field_value = field_value.strip("ADDIN ZOTERO_ITEM CSL_CITATION").strip()
                    field_value_json = loads(field_value)

                    # find the year string in the citation.
                    while oRangeFind.Execute("[0-9]{4}") and oRange.InRange(field.Result):
                        matchedText = oRange.Text
                        citations_list = field_value_json["citationItems"]

                        # loop each citation in a citation group
                        for _citation in citations_list:
                            citation_year = _citation["itemData"]["issued"]["date-parts"][0][0]

                            # find the right information
                            if citation_year == matchedText:
                                authors_list = _citation["itemData"]["author"]

                                if "language" not in _citation["itemData"]:
                                    # assume the default language is Chinese
                                    etal_text = "等"
                                    is_cn_language = False
                                else:
                                    language: str = _citation["itemData"]["language"]
                                    if language.lower() == "en":
                                        etal_text = "etal"
                                        is_cn_language = False
                                    else:
                                        etal_text = "等"
                                        is_cn_language = True

                                bmtext = _generate_bookmark_id(authors_list, etal_text, citation_year, etal_number, is_cn_language)

                                docx_obj.Hyperlinks.Add(Anchor=oRange, Address="", SubAddress=bmtext, ScreenTip="",
                                                        TextToDisplay="")
                                oRange.Collapse(0)


if __name__ == '__main__':
    word_file_path = r""
    new_file_path = r""
    error_flag = False

    # open word
    word = Dispatch("Word.Application")
    word.Visible = False
    docx = word.Documents.Open(word_file_path)

    try:
        create_bookmarks_for_zotero_literature(docx)
        create_hyperlinks_to_literature_bookmarks(docx, setColor=16711680, etal_number=3)
    except ZoteroCitationError:
        error_flag = True

    if not error_flag:
        docx.SaveAs(new_file_path)

    docx.Close(False)
    word.Quit()
