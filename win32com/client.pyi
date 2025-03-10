class Dispatch:
    def __init__(self, *args, **kwargs):
        self.Documents = None

    def Quit(self, *args, **kwargs):
        ...

class Documents:
    @classmethod
    def Open(cls, *args, **kwargs):
        ...
