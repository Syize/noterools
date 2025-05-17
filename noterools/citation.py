from json import loads
from os.path import basename

from .csl import CSLJson
from .error import AddHyperlinkError
from .hook import HookBase
from .utils import get_year_list, logger, replace_invalid_char, parse_color
from .word import Word


class CitationHyperlinkHook(HookBase):
    """
    Hook to add hyperlinks to citations.
    """

    def __init__(self, is_numbered=False, color=None, no_under_line=True, full_citation_hyperlink=False):
        super().__init__("CitationHyperlinkHook")
        self.is_numbered = is_numbered
        self.color = parse_color(color)  # Use parse_color
        self.no_under_line = no_under_line
        self.full_citation_hyperlink = full_citation_hyperlink

    def on_iterate(self, word_obj: Word, field):
        if "ADDIN ZOTERO_ITEM" not in field.Code.Text:
            return

        # We will change color after adding hyperlinks
        color_range = field.Result
        original_range = field.Result
        citation_text = original_range.Text

        # Handle numbered citations
        if self.is_numbered:
            temp_range = original_range.Duplicate
            temp_range.Collapse(1)
            range_find = temp_range.Find
            range_find.MatchWildcards = True

            # Find the number and add hyperlink
            while range_find.Execute("[0-9]{1,}") and temp_range.InRange(field.Result):
                bmtext = f"Ref_{temp_range.Text}"
                word_obj.add_hyperlink(bmtext, temp_range, no_under_line=self.no_under_line)
                temp_range.Collapse(0)
        
        # Handle author-year citations
        else:
            field_value: str = field.Code.Text.strip()
            field_value = field_value.strip("ADDIN ZOTERO_ITEM CSL_CITATION").strip()
            field_value_json = loads(field_value)
            citations_list = field_value_json["citationItems"]

            # Check if this is a multi-citation with semicolons
            if ';' in citation_text and '(' in citation_text and ')' in citation_text:
                # Process multi-citation (citations separated by semicolons)
                citation_content = citation_text.strip('()')
                citation_parts = [part.strip() for part in citation_content.split(';')]
                
                # Track position in the original text
                current_pos = 1  # Skip the opening parenthesis
                
                for part in citation_parts:
                    # Find this part in the original text, starting from current position
                    part_pos = citation_text.find(part, current_pos)
                    if part_pos == -1:
                        logger.warning(f"Could not locate citation part: '{part}'")
                        continue
                    
                    # Extract years from this citation part
                    part_years = get_year_list(part)
                    if not part_years:
                        logger.warning(f"No year found in citation part: '{part}'")
                        current_pos = part_pos + len(part)
                        continue
                    
                    # Match this citation part to its corresponding CSL citation item
                    matched = False
                    for _citation in citations_list:
                        item_key = basename(_citation["uris"][0])
                        csl_json = CSLJson(_citation["itemData"])
                        citation_year = str(csl_json.get_date().year)
                        language = csl_json.get_language(defaults="cn")
                        author_name = csl_json.get_author_names(language)[0]
                        
                        year_match = any(year[:4] in citation_year for year in part_years)
                        
                        # Check if this CSL item matches the current citation part
                        if (author_name in part and year_match) or (len(part) <= 7 and year_match):
                            # Create a duplicate range for this citation part
                            part_range = original_range.Duplicate
                            
                            if self.full_citation_hyperlink:
                                # Hyperlink the entire citation part
                                part_range.MoveStart(Unit=1, Count=part_pos)
                                part_range.MoveEnd(Unit=1, Count=-(len(citation_text) - (part_pos + len(part))))
                            else:
                                # Only hyperlink the year
                                year = part_years[0]  # Use the first year found
                                year_pos = part.find(year)
                                if year_pos == -1:
                                    logger.warning(f"Could not locate year in part: '{part}'")
                                    continue
                                    
                                absolute_year_pos = part_pos + year_pos
                                part_range.MoveStart(Unit=1, Count=absolute_year_pos)
                                part_range.MoveEnd(Unit=1, Count=-(len(citation_text) - (absolute_year_pos + len(year))))
                            
                            bmtext = f"Ref_{item_key}"
                            try:
                                word_obj.add_hyperlink(bmtext, part_range, no_under_line=self.no_under_line)
                                matched = True
                                break
                            except AddHyperlinkError:
                                logger.warning(f"Failed to add hyperlink for '{part}'")
                    
                    if not matched:
                        logger.warning(f"No matching reference found for citation part: '{part}'")
                    
                    # Move past this part for the next iteration
                    current_pos = part_pos + len(part)
                    
            else:
                # Process simple citation (single author-year or similar)
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

                    # Move range to the next year string
                    current_range = original_range.Duplicate
                    if is_first:
                        if not self.full_citation_hyperlink:
                            # Default: Only "Year" will have hyperlink
                            current_range.MoveStart(Unit=1, Count=len(authors_text))
                            current_range.MoveEnd(Unit=1, Count=-len(citation_text_left))
                        else:
                            # "Author, Date" will have hyperlink, but exclude the opening parenthesis
                            current_range.MoveStart(Unit=1, Count=1)  # Skip the opening parenthesis
                            current_range.MoveEnd(Unit=1, Count=-len(citation_text_left))
                        is_first = False
                    else:
                        # For subsequent years in a multi-year citation
                        year_pos = citation_text.find(_year, len(citation_text) - len(citation_text_left) - len(_year))
                        if year_pos != -1:
                            current_range.MoveStart(Unit=1, Count=year_pos)
                            current_range.MoveEnd(Unit=1, Count=-(len(citation_text) - (year_pos + len(_year))))
                        else:
                            # Fallback to original method if precise positioning fails
                            current_range.MoveEnd(Unit=1, Count=len(authors_text) + 5)
                            current_range.MoveStart(Unit=1, Count=len(authors_text) + 4)

                    is_add_hyperlink = False
                    for _citation in citations_list:
                        item_key = basename(_citation["uris"][0])
                        csl_json = CSLJson(_citation["itemData"])
                        citation_year = str(csl_json.get_date().year)
                        language = csl_json.get_language(defaults="cn")
                        author_name = csl_json.get_author_names(language)[0]

                        if multiple_article_for_one_author:
                            authors_text = last_authors_text

                        _year_without_character = _year[:4]

                        # Check match conditions
                        res1 = author_name in authors_text and _year_without_character in citation_year
                        res2 = replace_invalid_char(authors_text) == "" and _year_without_character in citation_year
                        res3 = citation_text_length <= 7

                        if res1 or res2 or res3:
                            bmtext = f"Ref_{item_key}"

                            try:
                                word_obj.add_hyperlink(bmtext, current_range, no_under_line=self.no_under_line)
                                is_add_hyperlink = True
                                break
                            except AddHyperlinkError:
                                is_add_hyperlink = False

                    if not is_add_hyperlink:
                        text = current_range.Text
                        current_range.MoveStart(Unit=1, Count=-20)
                        current_range.MoveEnd(Unit=1, Count=20)
                        logger.warning(f"Can't set hyperlinks for '{text}' in {current_range.Text}")
                        current_range.MoveStart(Unit=1, Count=20)
                        current_range.MoveEnd(Unit=1, Count=-20)

        # Apply color to the entire citation content (excluding parentheses)
        if self.color is not None:
            try:
                # Exclude "(" and ")" from color formatting
                color_range.MoveStart(Unit=1, Count=1)
                color_range.MoveEnd(Unit=1, Count=-1)
                color_range.Font.Color = self.color
            except Exception as e:
                logger.warning(f"Failed to apply color: {e}")


def add_citation_hyperlink_hook(word: Word, is_numbered=False, color=None, no_under_line=True, full_citation_hyperlink=False) -> CitationHyperlinkHook:
    """
    Register ``CitationHyperlinkHook``.

    :param word: ``noterools.word.Word`` object.
    :type word: Word
    :param is_numbered: If your citation is numbered. Defaults to False.
    :type is_numbered: bool
    :param color: Set font color. Accepts integer decimal value (e.g., 16711680 for blue), 
                 RGB string (e.g., "255, 0, 0" for red), or "word_auto" for automatic color.
                 You can look up the values at `VBA Documentation <https://learn.microsoft.com/en-us/office/vba/api/word.wdcolor>`_.
    :type color: Union[int, str, None]
    :param no_under_line: If remove the underline of hyperlinks. Defaults to True.
    :type no_under_line: bool
    :param full_citation_hyperlink: If True, the entire citation (author and year) will be hyperlinked for the first reference in multiple citations. For subsequent references in the same citation block, only the year will be hyperlinked due to technical limitations. Defaults to False (only year is hyperlinked).
    :type full_citation_hyperlink: bool
    :return: ``CitationHyperlinkHook`` instance.
    :rtype: CitationHyperlinkHook
    """
    citation_hyperlink_hook = CitationHyperlinkHook(is_numbered, color, no_under_line, full_citation_hyperlink)
    word.set_hook(citation_hyperlink_hook)

    return citation_hyperlink_hook


__all__ = ["CitationHyperlinkHook", "add_citation_hyperlink_hook"]
