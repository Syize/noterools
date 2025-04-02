class NoteroolsBasicError(Exception):
    """
    Basic exception class of Noterools.
    """
    pass


class ContextError(NoteroolsBasicError):
    """
    Not in Noterools context error.
    """
    pass


class AddBookmarkError(NoteroolsBasicError):
    """
    Can't add bookmark error.
    """
    pass


class AddHyperlinkError(NoteroolsBasicError):
    """
    Can't add hyperlink error.
    """
    pass


class HookTypeError(NoteroolsBasicError):
    """
    Unknown hook type.
    """
    pass


class ArticleNotFoundError(NoteroolsBasicError):
    """
    Article not found in zotero.
    """
    pass


class TitleNotFoundError(NoteroolsBasicError):
    """
    Article title not found.
    """
    pass


class AuthorNotFoundError(NoteroolsBasicError):
    """
    Article author not found.
    """
    pass


class ParamsError(NoteroolsBasicError):
    """
    Sme hooks may require the user to give at least one parameter.
    """
    pass


__all__ = ["NoteroolsBasicError", "AddBookmarkError", "AddHyperlinkError", "ContextError", "HookTypeError", "ArticleNotFoundError", "TitleNotFoundError", "AuthorNotFoundError",
           "ParamsError"]
