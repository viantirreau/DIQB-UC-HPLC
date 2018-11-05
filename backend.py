from PyQt5.QtCore import pyqtSignal, QObject, pyqtSlot
import urllib.parse
import os


class PDFToExcel(QObject):
    front_add_path = pyqtSignal(str)
    front_remove_path = pyqtSignal(str)

    def __init__(self, front_obj):
        super().__init__()
        print("Backend inicializado")
        self.names_paths = {}
        self.front_add_path.connect(front_obj.add_path_to_list)
        self.front_remove_path.connect(front_obj.drop_path_from_list)

    @pyqtSlot(str)
    def add_paths_drag_n_drop(self, paths):
        print("Back DnD: Me llamaron")
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
            if os.path.splitext(path_final)[1] == ".pdf":
                self.add_paths(path_final)

    @pyqtSlot(str)
    def add_paths(self, path):
        print("Back AddP: Me llamaron")
        display_name = os.path.basename(path)
        self.names_paths[display_name] = path
        self.front_add_path.emit(display_name)

    @pyqtSlot(str)
    def remove_paths(self, path):
        self.names_paths.pop(path, None)

    def test(self, *args):
        print("Funciona!")
