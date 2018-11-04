from PyQt5.QtCore import Qt
from PyQt5.QtGui import (QPalette, QPixmap, qRgba, QStandardItem,
                         QStandardItemModel)
from PyQt5.QtWidgets import QApplication, QLabel, QWidget, QPushButton, \
    QFileDialog, QVBoxLayout, QHBoxLayout, QListView
import os
import urllib.parse


class Drop(QWidget):
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
        for file in files.split('\n'):
            if not file:
                continue
            raw_path = urllib.parse.urlparse(file).path
            decoded = urllib.parse.unquote(raw_path)
            os_path = os.path.normpath(decoded)
            path_final = os_path
            drive = os.path.splitdrive(path_final[1:])[0]
            if drive:
                path_final = path_final[1:]
            print(path_final, os.path.isfile(path_final),
                  os.path.splitext(path_final)[1] == ".pdf")

        # if event.mimeData().hasFormat('application/pdf'):
        #     print("PDF")
        #     mime = event.mimeData()
        #     item_data = mime.data('application/pdf')
        #     data_stream = QDataStream(item_data, QIODevice.ReadOnly)
        #
        #     text = QByteArray()
        #     offset = QPoint()
        #     data_stream >> text >> offset
        #
        #     if event.source() in self.children():
        #         event.setDropAction(Qt.MoveAction)
        #         event.accept()
        #     else:
        #         event.acceptProposedAction()
        # else:
        #     event.ignore()

    def get_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Elegir uno o más archivos PDF",
            "",
            "Archivos PDF (*.pdf)")
        if files:
            print(files)

    def remove_checked(self):
        model = self.list.model()
        pos = 0
        while pos < model.rowCount():
            item = model.item(pos)
            if item.checkState() == Qt.Checked:
                model.removeRow(pos)
            else:
                pos += 1


if __name__ == '__main__':
    import sys

    app = QApplication(sys.argv)
    window = Drop()
    window.show()
    sys.exit(app.exec_())
