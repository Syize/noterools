from datetime import datetime
from json import loads
from os.path import basename
from typing import Union

from .error import TitleNotFoundError, AuthorNotFoundError
from .hook import HookBase, HOOKTYPE
from .utils import logger
from .word import Word


class CSLJson(dict):
    """
    A class to parse CSL JSON.
    """
    def __init__(self, csl_json: Union[str, dict]):
        """
        CSLJson provide some convenient methods to get article's information from CSL JSON.

        :param csl_json: A CSL JSON string or CSL JSON dict.
        :type csl_json: str | dict
        """
        if isinstance(csl_json, str):
            csl_json = loads(csl_json)

        super().__init__(csl_json)

    def __getitem__(self, item: str):
        return super().__getitem__(item)

    def get_type(self) -> str:
        if "type" not in self:
            return "article-journal"
        else:
            return self["type"]

    def get_title(self) -> str:
        """
        Get article's title.

        :return: Article's title.
        :rtype: str
        """
        if "title" not in self:
            logger.error(f"'title' not found in CSL json, please check your citations.")
            raise TitleNotFoundError(f"'title' not found in CSL json, please check your citations.")

        else:
            return self["title"]

    def get_container_title(self) -> str:
        """
        Get article's container title.

        :return: Article's container title. "" is returned if ``container-title`` not in CSL JSON.
        :rtype: str
        """
        if "container-title" not in self:
            return ""

        else:
            return self["container-title"]

    def get_publisher(self) -> str:
        """
        Get article's publisher name.

        :return: Article's publisher name. "" is returned if ``publisher`` not in CSL JSON.
        :rtype: str
        """
        if "publisher" not in self:
            return ""

        else:
            return self["publisher"]

    def get_language(self, defaults="en") -> str:
        """
        Get article's language in lower case.

        :param defaults: Specify the language when ``language`` not in CSL JSON.
        :type defaults: str
        :return: Article's language in lower case.
        :rtype: str
        """
        if "language" not in self:
            return defaults
        else:
            return self["language"].lower()

    def get_author_names(self, language="en") -> list[str]:
        """
        Get all authors' name.

        :param language: If ``cn``, return ``family name + given name``, otherwise only return ``family name``.
        :type language: str
        :return: All authors' name in a list.
        :rtype: list
        """
        if self.get_type() == "software":
            language = "en"

        if "author" in self:
            key_name = "author"

        elif "editor" in self:
            key_name = "editor"

        else:
            logger.error(f"'author' or 'editor' not found in CSL json, please check your citations.")
            raise AuthorNotFoundError(f"'author' or 'editor' not found in CSL json, please check your citations.")

        res = []
        for author in self[key_name]:
            if "family" in author:
                if language == "cn":
                    res.append(f"{author['family']}{author['given']}")
                else:
                    res.append(author["family"])
            else:
                res.append(author["literal"])

        return res

    def get_date(self) -> datetime:
        """
        Get article's date.

        :return: Article's date. "1900-01-01" is used if ``issued`` not in CSL JSON.
        :rtype: datetime
        """
        if "issued" not in self:
            return datetime(1900, 1, 1)
        else:
            date_list = self["issued"]["date-parts"][0]
            date_list = [int(x) for x in date_list]
            while len(date_list) < 3:
                date_list.append(1)
            return datetime(*tuple(date_list))


class GetCSLJsonHook(HookBase):
    """
    Parse the CSL JSON in Zotero's citations.
    """
    def __init__(self):
        super().__init__("GetCSLJsonHook")
        self.csl_json_dict: dict[str, CSLJson] = {}

    def get_csl_jsons(self) -> dict[str, CSLJson]:
        """
        Get all CSL JSON data.

        :return: A dict which key is Zotero's item id and value is a CSLJson object.
        :rtype: dict
        """
        return self.csl_json_dict

    def on_iterate(self, word, field):
        if "ADDIN ZOTERO_ITEM" not in field.Code.Text:
            return

        # convert string to JSON string.
        field_value: str = field.Code.Text.strip()
        field_value = field_value.strip("ADDIN ZOTERO_ITEM CSL_CITATION").strip()
        field_value_json = loads(field_value)
        citations_list = field_value_json["citationItems"]

        for _citation in citations_list:
            item_id = basename(_citation["uris"][0])

            if item_id not in self.csl_json_dict:
                self.csl_json_dict[item_id] = CSLJson(_citation["itemData"])


def add_get_csl_json_hook(word: Word) -> GetCSLJsonHook:
    """
    Register ``GetCSLJsonHook``.

    :param word: ``noterools.word.Word`` object.
    :type word: Word
    :return: ``GetCSLJsonHook`` instance.
    :rtype: GetCSLJsonHook
    """
    get_csl_json_hook = GetCSLJsonHook()
    word.set_hook(get_csl_json_hook, HOOKTYPE.IN_ITERATE)

    return get_csl_json_hook


__all__ = ["GetCSLJsonHook", "CSLJson", "add_get_csl_json_hook"]
