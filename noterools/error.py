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


class ZoteroCitationError(Exception):
    """
    A custom exception class
    """
    pass


class HookTypeError(NoteroolsBasicError):
    """
    Unknown hook type.
    """
    pass


__all__ = ["ZoteroCitationError", "NoteroolsBasicError", "AddBookmarkError", "AddHyperlinkError", "ContextError", "HookTypeError"]
