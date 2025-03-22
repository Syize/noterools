from .hook import HookBase
from .word import Word


class CrossRefStyleHook(HookBase):
    """
    Set style of cross-reference.
    """
    def __init__(self, color: int = None, bold=False, key_word: str = None):
        if key_word is None:
            self.key_word = ""
        else:
            self.key_word = key_word
        super().__init__(f"CrossRefStyleHook: {self.key_word}")
        self.color = color
        self.bold = bold

    def on_iterate(self, word, field):
        if "REF _Ref" in field.Code.Text and self.key_word in field.Result.Text:
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


def add_cross_ref_style_hook(word_obj: Word, color: int = None, bold=False, key_word: list[str] = None):
    """
    Set font style of the cross-reference.

    :param word_obj:
    :type word_obj:
    :param color:
    :type color:
    :param bold:
    :type bold:
    :param key_word:
    :type key_word:
    :return:
    :rtype:
    """
    if isinstance(key_word, list):
        for _key in key_word:
            word_obj.set_hook(CrossRefStyleHook(color, bold, str(_key)))


__all__ = ["add_cross_ref_style_hook", "CrossRefStyleHook"]
