from PyQt5.QtCore import Qt, pyqtSignal, pyqtSlot
from PyQt5.QtGui import (QPalette, QStandardItem, QStandardItemModel, QColor,
                         QIcon)
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, \
    QFileDialog, QVBoxLayout, QHBoxLayout, QListView, QLabel
import backend
import sys


class CustomStandardItem(QStandardItem):
    def __init__(self, path_str):
        super().__init__(path_str)
        self.file_name = path_str


class Drop(QWidget):
    add_path_signal = pyqtSignal(str)
    add_drag_n_drop_path_signal = pyqtSignal(str)
    remove_path_signal = pyqtSignal(str)
    export_all_signal = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)

        # Window creation
        bg = self.palette()
        bg.setColor(QPalette.Window, Qt.white)
        self.setPalette(bg)
        self.setFixedSize(280, 350)
        self.setWindowTitle("HPLC a Excel")
        self.setAcceptDrops(True)

        # Button creation
        self.btn_load_files = QPushButton("Cargar archivos", self)
        self.btn_export = QPushButton("Exportar todos a Excel", self)
        self.btn_remove_checked = QPushButton("Descartar seleccionados", self)

        # Label creation
        self.copyright = QLabel("DIQB UC - Víctor Tirreau - 2018", self)
        self.copyright.setStyleSheet("color: #444444;")
        self.copyright.setAlignment(Qt.AlignCenter)

        # Button connection
        self.btn_load_files.clicked.connect(self.get_files)
        self.btn_remove_checked.clicked.connect(self.remove_checked)
        self.btn_export.clicked.connect(self.open_export_dialog)

        # Signal connection
        self.back = backend.PDFToExcel(self)
        self.add_path_signal.connect(self.back.add_paths)
        self.remove_path_signal.connect(self.back.remove_paths)
        self.add_drag_n_drop_path_signal.connect(
            self.back.add_paths_drag_n_drop)
        self.export_all_signal.connect(self.back.export_pdf_to_excel)

        # List
        self.list = QListView(self)
        self.model = QStandardItemModel(self.list)
        self.list.setModel(self.model)

        # Layout
        lower_h_box = QHBoxLayout()
        lower_h_box.addWidget(self.btn_remove_checked)
        lower_h_box.addStretch(10)
        lower_h_box.addWidget(self.btn_export)
        v_box = QVBoxLayout()
        v_box.addWidget(self.btn_load_files)
        v_box.addWidget(self.list)
        v_box.addLayout(lower_h_box)
        v_box.addWidget(self.copyright)
        h_box = QHBoxLayout()
        h_box.addStretch()
        h_box.addLayout(v_box)
        h_box.addStretch()
        self.setLayout(h_box)

    def dragEnterEvent(self, event):
        if event.mimeData().hasText():
            event.acceptProposedAction()
        else:
            event.ignore()

    dragMoveEvent = dragEnterEvent

    def dropEvent(self, event):
        files = event.mimeData().text()
        self.add_drag_n_drop_path_signal.emit(files)

    def get_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Elegir uno o más archivos PDF",
            "",
            "Archivos PDF (*.pdf)")
        if files:
            for file in files:
                self.add_path_signal.emit(file)

    def remove_checked(self):
        model = self.list.model()
        to_be_removed = [model.item(i) for i in range(model.rowCount()) if
                         model.item(i).checkState() == Qt.Checked]
        for path in to_be_removed:
            self.remove_path_signal.emit(path.file_name)

    def add_path_to_list(self, path_str):
        model = self.list.model()
        item = CustomStandardItem(path_str)
        item.setCheckable(True)
        item.setToolTip("Aún no procesado")
        item.setEditable(False)
        model.appendRow(item)

    def drop_path_from_list(self, path_str):
        model = self.list.model()
        pos = 0
        while pos < model.rowCount():
            item = model.item(pos)
            if item.file_name == path_str:
                model.removeRow(pos)
                break
            pos += 1

    def open_export_dialog(self):
        path = QFileDialog.getExistingDirectory(self,
                                                "Elegir carpeta para guardar "
                                                "Excel's")
        self.export_all_signal.emit(path)

    @pyqtSlot(tuple)
    def change_color_finished(self, args):
        file_name, res = args
        model = self.list.model()
        pos = 0
        while pos < model.rowCount():
            item = model.item(pos)
            if item.file_name == file_name:
                if res == 0:
                    item.setBackground(QColor(200, 250, 200))
                    item.setToolTip(f"Procesado correcto - {file_name}")
                    text = item.file_name
                    text = text[:34] + "..." if len(text) > 37 else text
                    text = "OK   " + text.ljust(38)
                    item.setText(text)
                else:
                    item.setBackground(QColor(250, 150, 150))
                    if res == 1:
                        item.setToolTip(
                            f"ERROR: El archivo {file_name} no tiene "
                            f"estándares")
                        text = item.file_name
                        text = text[:17] + "..." if len(text) > 21 else text
                        text += "  no tiene estándares"
                        item.setText(text)
                    if res == 2:
                        item.setToolTip(
                            f"ERROR: El archivo {file_name} está abierto por"
                            f" otro programa")
                        text = item.file_name
                        text = text[:17] + "..." if len(text) > 21 else text
                        text += "  está en uso por otro programa"
                        item.setText(text)
                    if res == 3:
                        item.setToolTip(
                            f"ERROR: El archivo {file_name} dio un error "
                            f"desconocido")
                        text = item.file_name
                        text = text[:17] + "..." if len(text) > 21 else text
                        text += "  error desconocido"
                        item.setText(text)
                break
            pos += 1

    @pyqtSlot(str)
    def change_color_started(self, file_name):
        model = self.list.model()
        pos = 0
        while pos < model.rowCount():
            item = model.item(pos)
            if item.text() == file_name:
                item.setBackground(QColor(200, 200, 250))
                item.setToolTip(f"Procesando {file_name}")
                text = item.file_name
                text = text[:25] + "..." if len(text) > 28 else text
                text += "  procesando"
                item.setText(text)
                break
            pos += 1


if __name__ == '__main__':
    def hook(type, value, traceback):
        print(type)
        print(traceback)


    sys.__excepthook__ = hook

    app = QApplication([])
    app.setWindowIcon(QIcon("ui/Logo.png"))
    window = Drop()
    window.show()
    sys.exit(app.exec_())
