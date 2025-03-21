from enum import Enum


class HookBase:
    """
    Base class for all hooks.

    """
    def __init__(self, name: str):
        self.name = name

    def before_iterate(self, word):
        """
        This method will be called before iteration.

        :return:
        :rtype:
        """
        pass

    def on_iterate(self, word, field):
        """
        This method will be called each iteration.

        :param word:
        :type word:
        :param field:
        :type field:
        :return:
        :rtype:
        """
        raise NotImplementedError("Child class must implement this method")

    def after_iterate(self, word):
        """
        This method will be called after iteration.

        :return:
        :rtype:
        """
        pass


class HOOKTYPE(Enum):
    """
    Hook types.
    """
    BEFORE_ITERATE = 0
    IN_ITERATE = 1
    AFTER_ITERATE = 2


__all__ = ["HookBase", "HOOKTYPE"]
