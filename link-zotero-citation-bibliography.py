import re
from json import loads

import pywintypes
from rich.progress import Progress
from win32com.client import Dispatch, CDispatch


class ZoteroCitationError(Exception):
    pass


def multipleReplace(text: str):
    string_list = [":", ";", ".", ",", "：", "；", "。", "，", "'", "’", " ", "-"]
    for s in string_list:
        text = text.replace(s, "")

    return text


def getYear(text: str, extra_pattern: str):
    pattern = r'\b(19\d{2}|20\d{2})[abcde]?\b' + extra_pattern  # 匹配 1900-2099 之间的四位年份
    res = re.findall(pattern, text)
    if len(res) < 1:
        print(f"Cannot get year: {text}")
        res_year = ""
    else:
        res_year = res[0]
    return res_year


def getAuthors(text: str, year_str: str):
    if "View Profile" in text:
        print(text)
    text = text.split(year_str)[0].strip()
    text = text.replace(",", "_").replace(" ", "").strip("_").strip(".")
    return text


def insertZoteroLiteratureBookmarks(docx_obj, isNumbered=False):
    total = len(list(docx_obj.Fields))

    with Progress() as progress:
        pid = progress.add_task(f"[red]Adding bookmarks..[red]", total=total)

        for field in docx_obj.Fields:

            progress.advance(pid, advance=1)

            if "ADDIN ZOTERO_BIBL" in field.Code.Text:
                oRange = field.Result

                for oBookMark in oRange.Bookmarks:
                    oBookMark.Delete()

                iCount = 1
                for oPara in oRange.Paragraphs:
                    oRangePara = oPara.Range
                    bmRange = oRangePara

                    if isNumbered:
                        bmtext = f"Ref_{iCount}"
                    else:
                        oYear = getYear(oRangePara.Text, r"[;；。，：:.,\)]")
                        if oYear == "":
                            print("GetYear error")
                            print(oRangePara.Text)
                            raise ZoteroCitationError
                        oAuthors = getAuthors(oRangePara.Text, oYear)
                        bmtext = "Ref_" + oAuthors + "_" + oYear
                        bmtext = multipleReplace(bmtext)

                    bmRange.MoveEnd(1, -1)
                    try:
                        docx_obj.Bookmarks.Add(Name=bmtext, Range=bmRange)
                    except pywintypes.com_error:
                        print("无法创建书签")
                        print(bmtext)
                        raise ZoteroCitationError
                    bmRange.Collapse(0)


def insertCitationToZoteroLiteratue(docx_obj, isNumbered=False):
    total = len(list(docx_obj.Fields))

    with Progress() as progress:
        pid = progress.add_task(f"[red]Adding hyperlinks..[red]", total=total)

        for field in docx_obj.Fields:

            progress.advance(pid, advance=1)

            if "ADDIN ZOTERO_ITEM" in field.Code.Text:
                oRange = field.Result
                oRangeFind = oRange.Find
                oRangeFind.MatchWildcards = True

                if isNumbered:
                    oRange.Collapse(1)

                    while oRangeFind.Execute("[0-9]{1,}") and oRange.InRange(field.Result):
                        bmtext = f"Ref_{oRange.Text}"
                        docx_obj.Hyperlinks.Add(Anchor=oRange, Address="", SubAddress=bmtext, ScreenTip="",
                                                TextToDisplay="")
                        oRange.Collapse(0)

                else:

                    field_value: str = field.Code.Text.strip()
                    field_value = field_value.strip("ADDIN ZOTERO_ITEM CSL_CITATION").strip()
                    field_value_json = loads(field_value)

                    while oRangeFind.Execute("[0-9]{4}") and oRange.InRange(field.Result):
                        matchedText = oRange.Text
                        citations_list = field_value_json["citationItems"]

                        for _citation in citations_list:
                            citation_year = _citation["itemData"]["issued"]["date-parts"][0][0]

                            if citation_year == matchedText:
                                authors_list = _citation["itemData"]["author"]

                                if "language" not in _citation["itemData"]:
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

                                if len(authors_list) > 3:

                                    _text = ""
                                    for _author in authors_list[:3]:
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
                                    _text += f"_{etal_text}"

                                    _text = multipleReplace(_text)
                                    bmtext = f"Ref_{_text}_{citation_year}"

                                else:
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

                                    _text = multipleReplace(_text)
                                    bmtext = f"Ref_{_text}_{citation_year}"

                                # oRange.MoveEnd(Unit=1, Count=-1)
                                docx_obj.Hyperlinks.Add(Anchor=oRange, Address="", SubAddress=bmtext, ScreenTip="",
                                                        TextToDisplay="")
                                oRange.Collapse(0)


if __name__ == '__main__':
    word_file_path = r""
    new_file_path = r""
    error_flag = False

    # open word
    word: CDispatch = Dispatch("Word.Application")
    word.Visible = False
    docx = word.Documents.Open(word_file_path)

    try:
        insertZoteroLiteratureBookmarks(docx)
        insertCitationToZoteroLiteratue(docx)
    except ZoteroCitationError:
        error_flag = True

    if not error_flag:
        docx.SaveAs(new_file_path)

    docx.Close(False)
    word.Quit()
