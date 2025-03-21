import pywintypes
from rich.progress import Progress

from .error import ZoteroCitationError
from .utils import logger, get_citations_info


def add_bookmarks_to_bibliography(docx_obj, isNumbered=False, set_container_title_italic=True):
    """
    Add bookmarks to bibliographies.

    :param docx_obj: Docx object opened by pywin32.
    :param isNumbered: If the citation format is numbered.
    :param set_container_title_italic: If set the container-title and publisher of Chinese paper to Italic.
    :return:
    """
    title_item_key_dict = get_citations_info(docx_obj)
    title_container_title_tuple = [
        (
            title, title_item_key_dict[title]["container_title"], title_item_key_dict[title]["author"], title_item_key_dict[title]["publisher"],
            title_item_key_dict[title]["language"],
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
                if set_container_title_italic and bib_language == "cn":

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
                try:
                    docx_obj.Bookmarks.Add(Name=bmtext, Range=bmRange)
                except pywintypes.com_error:    # type: ignore
                    logger.error(f"Cannot add bookmarks: {bmtext}")
                    raise ZoteroCitationError
                bmRange.Collapse(0)


__all__ = ["add_bookmarks_to_bibliography"]
