from threading import Event
from PySide6.QtCore import QThread, Signal
from converter import convert


class ConvertThread(QThread):
    progress = Signal(int)
    success = Signal()
    error = Signal(str)
    cancelled = Signal()

    def __init__(self, input_path, mapping_path, template_path, output_path):
        super().__init__()

        self.input_path = input_path
        self.mapping_path = mapping_path
        self.template_path = template_path
        self.output_path = output_path

        self._cancel_event = Event()

    def stop(self):
        self._cancel_event.set()

    def run(self):
        try:
            completed = convert(
                input_path=self.input_path,
                mapping_path=self.mapping_path,
                template_path=self.template_path,
                output_path=self.output_path,
                progress_callback=self.progress.emit,
                cancel_flag=self._cancel_event.is_set,
            )

            if not completed or self._cancel_event.is_set():
                self.cancelled.emit()
                return

            self.progress.emit(100)
            self.success.emit()

        except Exception as error:
            self.error.emit(str(error))