from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import (QPalette, QPixmap, qRgba, QStandardItem,
                         QStandardItemModel)
from PyQt5.QtWidgets import QApplication, QLabel, QWidget, QPushButton, \
    QFileDialog, QVBoxLayout, QHBoxLayout, QListView
import backend


class Drop(QWidget):
    add_path_signal = pyqtSignal(str)
    add_drag_n_drop_path_signal = pyqtSignal(str)
    remove_path_signal = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)

        bg = self.palette()
        bg.setColor(QPalette.Window, Qt.white)
        self.setPalette(bg)
        h_box = QHBoxLayout()
        h_box.addStretch()
        v_box = QVBoxLayout()

        # self.setMinimumSize(280, 350)
        self.setFixedSize(280, 350)
        self.setWindowTitle("HPLC a Excel")
        self.setAcceptDrops(True)
        self.btn_load_files = QPushButton("Cargar archivos", self)
        lower_h_box = QHBoxLayout()
        self.btn_export = QPushButton("Exportar a Excel", self)
        self.btn_remove_checked = QPushButton("Descartar seleccionados", self)
        lower_h_box.addWidget(self.btn_remove_checked)
        lower_h_box.addStretch(10)
        lower_h_box.addWidget(self.btn_export)

        self.btn_load_files.clicked.connect(self.get_files)
        self.btn_remove_checked.clicked.connect(self.remove_checked)

        self.back = backend.PDFToExcel(self)
        self.add_path_signal.connect(self.back.add_paths)
        self.add_drag_n_drop_path_signal.connect(
            self.back.add_paths_drag_n_drop)

        self.list = QListView(self)
        self.model = QStandardItemModel(self.list)
        v_box.addWidget(self.btn_load_files)
        v_box.addWidget(self.list)
        v_box.addLayout(lower_h_box)
        h_box.addLayout(v_box)
        h_box.addStretch()

        codes = [
            'LOAA-05379',
            'LOAA-04468',
            'LOAA-03553',
            'LOAA-02642',
            'LOAA-05731'
        ]

        for code in codes:
            item = QStandardItem(code)
            item.setCheckable(True)
            item.setEditable(False)
            self.model.appendRow(item)
        self.list.setModel(self.model)

        self.setLayout(h_box)

    def dragEnterEvent(self, event):
        if event.mimeData().hasText():
            event.acceptProposedAction()
        else:
            event.ignore()

    dragMoveEvent = dragEnterEvent

    def dropEvent(self, event):
        files = event.mimeData().text()
        self.test_signal.emit()
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
                print("Señal emitida desde Dialog", file)

    def remove_checked(self):
        model = self.list.model()
        pos = 0
        while pos < model.rowCount():
            item = model.item(pos)
            if item.checkState() == Qt.Checked:
                model.removeRow(pos)
                self.remove_path_signal.emit(item.text())
            else:
                pos += 1

    def add_path_to_list(self, path_str):
        model = self.list.model()
        item = QStandardItem(path_str)
        item.setCheckable(True)
        item.setEditable(False)
        model.appendRow(item)

    def drop_path_from_list(self, path_str):
        model = self.list.model()
        pos = 0
        while pos < model.rowCount():
            item = model.item(pos)
            if item.text() == path_str:
                model.removeRow(pos)
                break
            pos += 1


if __name__ == '__main__':
    import sys

    app = QApplication(sys.argv)
    window = Drop()
    window.show()
    sys.exit(app.exec_())
