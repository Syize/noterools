import pywintypes
from rich.progress import Progress

from .error import ZoteroCitationError
from .utils import logger, get_citations_info


def add_bookmarks_to_bibliography(docx_obj, isNumbered=False):
    """
    Add bookmarks to bibliographies.

    :param docx_obj: Docx object opened by pywin32.
    :param isNumbered: If the citation format is numbered.
    :return:
    """
    title_item_key_dict = get_citations_info(docx_obj)
    title_publisher_tuple = [
        (
            title, title_item_key_dict[title]["publisher"], title_item_key_dict[title]["author"]
        ) for title in title_item_key_dict
    ]

    logger.info(f"Find bibliographies in the word, it may take a few seconds...")
    # loop fields in docx
    for field in docx_obj.Fields:

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

                if isNumbered:
                    bmtext = f"Ref_{iCount}"
                    iCount += 1

                else:
                    text = oRangePara.Text
                    bib_title = ""

                    for index, _tuple in enumerate(title_publisher_tuple):
                        _title, _publisher, _author = _tuple
                        if _title in text and _publisher in text and _author in text and f"{_title} " not in text:
                            bib_title = _title
                            title_publisher_tuple.pop(index)
                            break

                    if bib_title == "":
                        logger.warning(f"Can't find the corresponding citation of bib: {text}, do you really cite it?")
                        continue

                    bib_item_key = title_item_key_dict.pop(bib_title)["item_key"]
                    bmtext = f"Ref_{bib_item_key}"

                bmRange.MoveEnd(1, -1)
                try:
                    docx_obj.Bookmarks.Add(Name=bmtext, Range=bmRange)
                except pywintypes.com_error:    # type: ignore
                    logger.error(f"Cannot add bookmarks: {bmtext}")
                    raise ZoteroCitationError
                bmRange.Collapse(0)


__all__ = ["add_bookmarks_to_bibliography"]
