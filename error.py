class ReadError(Exception):
    def __init__(self, errorinfo):
        super().__init__(self)
        self.errorinfo = errorinfo

    def __str__(self):
        return self.errorinfo


class GradeError(Exception):
    def __init__(self, errorinfo):
        super().__init__(self)
        self.errorinfo = errorinfo

    def __str__(self):
        return self.errorinfo


class ComprehensionError(Exception):
    def __init__(self, errorinfo):
        super().__init__(self)
        self.errorinfo = errorinfo

    def __str__(self):
        return self.errorinfo


class WeChatError(Exception):
    def __init__(self, errorinfo):
        super().__init__(self)
        self.errorinfo = errorinfo

    def __str__(self):
        return self.errorinfo


class ModelError(Exception):
    def __init__(self, errorinfo):
        super().__init__(self)
        self.errorinfo = errorinfo

    def __str__(self):
        return self.errorinfo
