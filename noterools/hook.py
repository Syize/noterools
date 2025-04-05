from enum import Enum

from .error import HookTypeError


class HOOKTYPE(Enum):
    """
    Hook types.
    """
    BEFORE_ITERATE = 0
    IN_ITERATE = 1
    AFTER_ITERATE = 2


class HookBase:
    """
    Base class for all hooks.
    """
    # singleton pattern
    _hook_instance = None
    # make sure attributes defined in ``HookBase`` are only initialized once.
    _hook_initialized = False

    def __init__(self, name: str):
        if self._hook_initialized:
            return

        # unique name for specific hook
        self.name = name
        # tell if the hook is registered
        self._registered_before_iterate = False
        self._registered_in_iterate = False
        self._registered_after_iterate = False

        self._hook_initialized = True

    def __new__(cls, *args, **kwargs):
        if cls._hook_instance is None:
            cls._hook_instance = super().__new__(cls)

        return cls._hook_instance

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

    def finish_register(self, hook_type: HOOKTYPE = HOOKTYPE.IN_ITERATE):
        """
        Set register flag to True.

        :param hook_type: Hook type. Defaults to ``IN_ITERATE``.
        :type hook_type: HOOKTYPE
        :return:
        :rtype:
        """
        if hook_type == HOOKTYPE.BEFORE_ITERATE:
            self._registered_before_iterate = True
        elif hook_type == HOOKTYPE.IN_ITERATE:
            self._registered_in_iterate = True
        elif hook_type == HOOKTYPE.AFTER_ITERATE:
            self._registered_after_iterate = True
        else:
            raise HookTypeError(f"Unknown hook type: {hook_type}.")

    def is_registered(self, hook_type: HOOKTYPE = HOOKTYPE.IN_ITERATE) -> bool:
        """
        Return the register flag.

        :param hook_type: Hook type. Defaults to ``IN_ITERATE``.
        :type hook_type: HOOKTYPE
        :return: Register flag.
        :rtype: bool
        """
        if hook_type == HOOKTYPE.BEFORE_ITERATE:
            return self._registered_before_iterate
        elif hook_type == HOOKTYPE.IN_ITERATE:
            return self._registered_in_iterate
        elif hook_type == HOOKTYPE.AFTER_ITERATE:
            return self._registered_after_iterate
        else:
            raise HookTypeError(f"Unknown hook type: {hook_type}.")


class ExtensionHookBase:
    """
    Base class for all extension hooks.
    """
    # singleton pattern
    _hook_instance = None
    # make sure attributes defined in ``ExtensionHookBase`` are only initialized once.
    _hook_initialized = False

    def __init__(self, name: str):
        if self._hook_initialized:
            return

        # unique name for specific hook
        self.name = name
        # tell if the hook is registered
        self._registered = False

        self._hook_initialized = True

    def __new__(cls, *args, **kwargs):
        if cls._hook_instance is None:
            cls._hook_instance = super().__new__(cls)

        return cls._hook_instance

    def on_iterate(self, word, word_range):
        """
        This method will be called each iteration.

        :param word:
        :type word:
        :param word_range:
        :type word_range:
        :return:
        :rtype:
        """
        raise NotImplementedError("Child class must implement this method")

    def finish_register(self):
        """
        Set register flag to True.

        :return:
        :rtype:
        """
        self._registered = True

    def is_registered(self) -> bool:
        """
        Return the register flag.

        :return: Register flag.
        :rtype: bool
        """
        return self._registered


__all__ = ["HookBase", "HOOKTYPE", "ExtensionHookBase"]
