from .hook import HookBase
from .utils import logger
from .word import Word


class CrossRefStyleHook(HookBase):
    """
    Set style of cross-reference.
    """
    def __init__(self, color: int = None, bold=False, key_word: list[str] = None):
        super().__init__(f"CrossRefStyleHook")

        if key_word is None:
            self.key_word = [""]

        else:
            if len(key_word) > 1 and "" in key_word:
                logger.warning("Found empty string in your keyword. It can cause noterools to change all of your cross references")
                key_word = [x for x in key_word if x != ""]
                logger.warning(f"All empty strings have been removed, new key word list is: {key_word}")

            self.key_word = key_word

        self.color = color
        self.bold = bold

    def on_iterate(self, word, field):
        if "REF _Ref" not in field.Code.Text:
            return

        for _key_word in self.key_word:
            if _key_word not in field.Result.Text:
                continue

            # update field code so it keeps settings
            field_code = field.Code.Text
            if r"\* MERGEFORMAT" not in field_code:
                # if you miss the white space at the last of code, Word can't recognize the field code.
                # Word is shit.
                field_code += r" \* MERGEFORMAT "
                field.Code.Text = field_code

            range_obj = field.Result
            if self.color is not None:
                range_obj.Font.Color = self.color
            range_obj.Font.Bold = self.bold

            range_obj = field.Code
            range_obj.MoveStart(Unit=1, Count=-1)
            range_obj.MoveEnd(Unit=1, Count=1)
            if self.color is not None:
                range_obj.Font.Color = self.color
            range_obj.Font.Bold = self.bold

        field.Update()


def add_cross_ref_style_hook(word: Word, color: int = 16711680, bold=False, key_word: list[str] = None) -> CrossRefStyleHook:
    """
    Set font style of the cross-reference.

    :param word: ``noterools.word.Word`` object.
    :type word: Word
    :param color: Set font color. Defaults to ``blue (16711680)``. You can look up the value at `VBA Documentation
                  <https://learn.microsoft.com/en-us/office/vba/api/word.wdcolor>`_.
    :type color: int
    :param bold: If you make font bold. Defaults to False.
    :type bold: bool
    :param key_word: A key word list. noterools only change the cross-ref's style which contains the key word you specify, for example, ``["Fig.", "Tab."]``.
                     noterools will change all cross-refs style if ``ket_word == None``.
    :type key_word: list
    :return:
    :rtype:
    """
    if key_word is None:
        logger.warning("Set style for all cross references because `key_word` is None")

    cross_ref_style_hook = CrossRefStyleHook(color, bold, key_word)
    word.set_hook(cross_ref_style_hook)

    return cross_ref_style_hook


__all__ = ["add_cross_ref_style_hook", "CrossRefStyleHook"]
