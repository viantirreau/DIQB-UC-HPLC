from PyQt5.QtCore import pyqtSignal, QObject, QThreadPool, QRunnable, \
    pyqtSlot, QThread
import urllib.parse
import os
from dict_to_xl import dict_to_xlsx


class PDFToExcel(QObject):
    front_add_path = pyqtSignal(str)
    front_remove_path = pyqtSignal(str)

    def __init__(self, front_obj):
        super().__init__()
        self.names_paths = {}
        self.front = front_obj
        self.front_add_path.connect(self.front.add_path_to_list)
        self.front_remove_path.connect(self.front.drop_path_from_list)

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

    def export_pdf_to_excel(self, args):
        output_path, include_od = args
        output_path = os.path.normpath(output_path)
        pool = QThreadPool()
        # pool.globalInstance().setMaxThreadCount(2)
        for front_name, input_path in self.names_paths.items():
            worker = Worker(dict_to_xlsx, front_name, input_path,
                            output_path, report_od=include_od)
            worker.signals.progress.connect(self.front.progress_started)
            worker.signals.result.connect(self.front.change_color_finished)
            pool.globalInstance().start(worker)  # .globalInstance() is the key!


class WorkerSignals(QObject):
    finished = pyqtSignal()
    error = pyqtSignal(tuple)
    result = pyqtSignal(tuple)
    progress = pyqtSignal(tuple)
    internal_progress = pyqtSignal(float)


class Worker(QRunnable):
    def __init__(self, task, name, *args, **kwargs):
        super().__init__()
        self.task = task
        self.front_name = name
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()
        self.signals.internal_progress.connect(self.report_progress)
        self.kwargs["sgn_progress"] = self.signals.internal_progress

    @pyqtSlot()
    def run(self):
        res = self.task(*self.args, **self.kwargs)
        if res == 0:
            self.signals.result.emit((self.front_name, 0))
        elif res == 1:
            self.signals.result.emit((self.front_name, 1))
        elif res == 2:
            self.signals.result.emit((self.front_name, 2))
        else:
            self.signals.result.emit((self.front_name, 3))

    def report_progress(self, progress):
        self.signals.progress.emit((self.front_name, progress))
