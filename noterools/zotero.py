from typing import Optional

from pyzotero.zotero import Zotero

from .error import ArticleNotFoundError
from .utils import logger

ZOTERO_CLIENT: Optional[Zotero] = None


def init_zotero_client(zotero_id: str, zotero_api_key: str, force=False):
    """
    Initial the client to Zotero.

    :param zotero_id: User ID.
    :param zotero_api_key: API key of Zotero.
    :param force: Force to re-initial Zotero client.
    :return:
    """
    global ZOTERO_CLIENT

    if ZOTERO_CLIENT is None or force:
        ZOTERO_CLIENT = Zotero(zotero_id, "user", zotero_api_key)


def search_article(itemKey: str) -> dict:
    """
    Use the itemKey to search the corresponding article in Zotero.

    :param itemKey: Zotero item key.
    :return: Article info.
    """
    global ZOTERO_CLIENT

    res = ZOTERO_CLIENT.items(itemKey=itemKey)
    if len(res) == 0:
        logger.error(f"Can't find the article which itemKey is {itemKey}")
        raise ArticleNotFoundError(f"Can't find the article which itemKey is {itemKey}")

    elif len(res) > 1:
        logger.warning(f"Find multiple articles which itemKey are same, use the first one")

    return res[0]


def get_author_list(item: dict) -> list[str]:
    """
    Get full names of authors from the item.

    :param item: Item from Zotero.
    :return: Authors' name list.
    """
    creators = item["data"]["creators"]
    name_list = []

    for creator in creators:
        if creator["creatorType"] != "author":
            continue

        if "name" in creator:
            name_list.append(creator["name"])

        else:
            first_name = creator["firstName"]

            if 65 <= ord(first_name[0]) <= 90 or 97 <= ord(first_name[0]) <= 122:
                # English name
                name_list.append(creator["firstName"] + creator["lastName"])

            else:
                # Chinese name
                name_list.append(creator["lastName"] + creator["firstName"])

    return name_list


__all__ = ["init_zotero_client", "search_article", "get_author_list"]
