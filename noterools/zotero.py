from typing import Optional

import httpx
from pyzotero.zotero import Zotero
from pyzotero.zotero_errors import ResourceNotFoundError

from .error import ZoteroNotInitError, ZoteroConnectError
from .utils import logger

ZOTERO_CLIENT: Optional[Zotero] = None


def zotero_check_initialized(func_name: str):
    """
    Check is Zotero client is initialized.
    An error will be raised if not.

    :param func_name: Name of the module which calls this function.
    :type func_name: str
    :return:
    :rtype:
    """
    global ZOTERO_CLIENT
    
    if ZOTERO_CLIENT is None:
        logger.error(f"You need to initialize the Zotero client with function `zotero_init_client`, cause '{func_name}' needs it.")
        raise ZoteroNotInitError(f"You need to initialize the Zotero client with function `zotero_init_client`, cause '{func_name}' needs it.")


def zotero_init_client(zotero_id: str = "0", zotero_api_key: Optional[str] = None, local=False, force=False):
    """
    Initial the client to Zotero.

    :param zotero_id: Zotero ID.
    :type zotero_id: str
    :param zotero_api_key: API key of Zotero.
    :type zotero_api_key: str
    :param local: Use local Zotero api or not?
    :type local: bool
    :param force: Force to re-initialize the Zotero client.
    :type force: bool
    :return:
    """
    global ZOTERO_CLIENT

    if ZOTERO_CLIENT is None or force:
        ZOTERO_CLIENT = Zotero(zotero_id, "user", zotero_api_key, local=local)


def zotero_query(item_id: str) -> dict:
    """
    Query item information from the Zotero.

    :param item_id: Zotero item ID.
    :type item_id: str
    :return: Item information.
    :rtype: dict
    """
    global ZOTERO_CLIENT

    try:
        return ZOTERO_CLIENT.item(item_id, format="json")
    except ResourceNotFoundError:
        return {"data": {}}
    except httpx.ConnectError:
        logger.error(f"Can't communicate with Zotero.")
        raise ZoteroConnectError(f"Can't communicate with Zotero.")


def zotero_query_pages(item_id: str) -> str:
    """
    Query item pages from the Zotero.

    :param item_id: Zotero item ID.
    :type item_id: Item information.
    :return: Item page range.
    :rtype: str
    """
    item_info = zotero_query(item_id)["data"]

    if "pages" in item_info:
        return item_info["pages"]
    else:
        return ""


__all__ = ["zotero_init_client", "zotero_query", "zotero_query_pages", "zotero_check_initialized"]
