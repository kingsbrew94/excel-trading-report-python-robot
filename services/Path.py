
class Path:
    _path = None

    @staticmethod
    def get_path():
        return Path._path

    @staticmethod
    def set_path(path):
        Path._path = path
