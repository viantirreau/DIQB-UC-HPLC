from PyQt5.QtCore import pyqtSignal, QObject, pyqtSlot
import urllib.parse
import os


class PDFToExcel(QObject):
    front_add_path = pyqtSignal(str)
    front_remove_path = pyqtSignal(str)

    def __init__(self, front_obj):
        super().__init__()
        self.names_paths = {}
        self.front_add_path.connect(front_obj.add_path_to_list)
        self.front_remove_path.connect(front_obj.drop_path_from_list)

    def add_paths_drag_n_drop(self, paths):
        for file in paths.split('\n'):
            if not file:
                continue
            raw_path = urllib.parse.urlparse(file).path
            decoded = urllib.parse.unquote(raw_path)
            os_path = os.path.normpath(decoded)
            path_final = os_path
            drive = os.path.splitdrive(path_final[1:])[0]
            if drive:
                path_final = path_final[1:]
            if not os.path.isfile(path_final):
                continue
            if path_final in self.names_paths.values():
                continue
            if os.path.splitext(path_final)[1] == ".pdf":
                self.add_paths(path_final)

    def add_paths(self, path):
        path = os.path.normpath(path)
        if path in self.names_paths.values():
            return None
        display_name = os.path.basename(path)
        self.names_paths[display_name] = path
        self.front_add_path.emit(display_name)

    def remove_paths(self, path):
        self.names_paths.pop(path, None)
        self.front_remove_path.emit(path)
