import sys
import os
from PySide6.QtWidgets import (QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel,
    QLineEdit, QMessageBox, QProgressBar, QHBoxLayout, QApplication, QProgressDialog)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QIcon, QPixmap
from worker import ConvertThread
from help import HelpDialog
import subprocess
from updater import UpdateCheckThread, UpdateDownloadThread
from version import APP_VERSION

def resource_path(relative_path):
    if hasattr(sys, "_MEIPASS"):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, relative_path)

#Config auto load
mapping_path = resource_path("config/mapping.json")
template_path = resource_path("config/template.xlsx")

# GUI
class ConverterUI(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowIcon(QIcon(resource_path("images/logo.png")))
        self.setWindowTitle("Databank Export Converter")
        self.setAcceptDrops(True)

        self.is_running = False

        layout = QVBoxLayout()

        #logo
        logo_label = QLabel()
        pixmap = QPixmap(resource_path("images/logo.png"))
        pixmap = pixmap.scaled(
            300,
            300,
            Qt.KeepAspectRatio,
            Qt.SmoothTransformation
        )

        logo_label.setPixmap(pixmap)
        logo_label.setAlignment(Qt.AlignLeft)
        layout.addWidget(logo_label)

        #Title
        title = QLabel("Databank Export Converter")
        title.setAlignment(Qt.AlignLeft)
        title.setStyleSheet("""
        font-size: 18pt;
        font-weight: bold
        """)

        layout.addWidget(title)
        layout.addSpacing(10)

        # Load Databank Export
        self.input_path = QLineEdit()
        self.input_path.setPlaceholderText("Load Databank export .csv file")

        layout.addWidget(QLabel("Input"))
        layout.addWidget(self.input_path)

        browse_csv = QPushButton("Load CSV File")
        browse_csv.clicked.connect(self.pick_csv)

        layout.addWidget(browse_csv)

        # Write output file
        self.output_path = QLineEdit()
        self.output_path.setPlaceholderText("Save output file")

        layout.addWidget(QLabel("Output File"))
        layout.addWidget(self.output_path)

        browse_output = QPushButton("Save Output As")
        browse_output.clicked.connect(self.pick_output)

        layout.addWidget(browse_output)

        #Buttons
        self.run_btn = QPushButton("Run Conversion")
        self.run_btn.clicked.connect(self.run_clicked)
        self.run_btn.setFixedHeight(50)

        self.help_btn = QPushButton("How to use")
        self.help_btn.clicked.connect(self.show_help)
        self.help_btn.setFixedHeight(50)
        layout.addSpacing(15)

        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        button_layout.addWidget(self.run_btn)
        button_layout.addWidget(self.help_btn)

        layout.addLayout(button_layout)

        layout.addSpacing(10)

        #Progress bar
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.hide()

        layout.addWidget(self.progress)

        self.setLayout(layout)

        QTimer.singleShot(1000, self.check_updates)

    #Check for updates
    def check_updates(self):
        self.update_check_thread = UpdateCheckThread()

        self.update_check_thread.update_available.connect(
            self.update_available
        )

        self.update_check_thread.no_update.connect(
            self.no_update
        )

        self.update_check_thread.error.connect(
            self.update_check_error
        )

        self.update_check_thread.start()

    def update_available(self, update_info):
        latest_ver = update_info["version"]
        release_notes = update_info["notes"]
        download_url = update_info["download_url"]

        up_msg = (
            f"A new version {latest_ver} of the converter is available.\n\n"
            f"Download and install now?"
        )

        ans = QMessageBox.question(
            self,
            "Update Available",
            up_msg,
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes
        )

        if ans == QMessageBox.Yes:
            self.download_update(download_url)

    def no_update(self):
        pass

    def update_check_error(self, err_msg):
        print(f"Update check failed: {err_msg}")

    def download_update(self, download_url):
        self.update_progress = QProgressDialog(
            "Downloading update...",
            None,
            0,
            100,
            self
        )

        self.update_progress.setWindowTitle("Downloading update")
        self.update_progress.setWindowModality(Qt.WindowModal)
        self.update_progress.setAutoClose(False)
        self.update_progress.setAutoReset(False)
        self.update_progress.setMinimumDuration(0)
        self.update_progress.setValue(0)

        self.update_download_thread = UpdateDownloadThread(
            download_url
        )

        self.update_download_thread.progress.connect(
            self.update_progress.setValue
        )

        self.update_download_thread.error.connect(
            self.on_update_download_error
        )

        self.update_download_thread.start()

    def install_update(self, installer_path):
        self.update_progress.setValue(100)
        self.update_progress.close()

        try:
            subprocess.Popen([installer_path])
        except Exception as error:
            QMessageBox.critical(
                self,
                "Update Error",
                "The update was downloaded but the installer "
                f"could not be started.\n\n{error}"
            )
            return

        QApplication.quit()

    def update_download_error(self, message):
        self.update_progress.close()

        QMessageBox.critical(
            self,
            "Update download failed",
            "The update could not be downloaded.\n\n"
            f"{message}"
        )

    #File pickers
    def pick_csv(self):
        input_file, _ = QFileDialog.getOpenFileName(
            self,
            "Load Databank Export file",
            "",
            "CSV Files (*.csv)"
        )

        if input_file:
            self.input_path.setText(input_file)
            self.output_path.clear()

    def pick_output(self):
        output_file, _ = QFileDialog.getSaveFileName(
            self,
            "Save Output File",
            "",
            "Excel Files (*.xlsx)"
        )

        if output_file:
            self.output_path.setText(output_file)

    #Drag & drop capability
    def dragEnterEvent(self, event):
        if not event.mimeData().hasUrls():
            event.ignore()
            return

        for url in event.mimeData().urls():
            input_file = url.toLocalFile()

            if input_file.lower().endswith(".csv"):
                event.acceptProposedAction()
                return

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            input_file = url.toLocalFile()

            if input_file.lower().endswith(".csv"):
                self.input_path.setText(input_file)
                self.output_path.clear()
                event.acceptProposedAction()
                return

        event.ignore()

    #Dynamic run/cancel button
    def run_clicked(self):
        if not self.is_running:
            self.start_conversion()
        else:
            self.cancel_conversion()

    #Start conversion
    def start_conversion(self):
        input_path = self.input_path.text().strip()
        output_path = self.output_path.text().strip()

        if not input_path:
            QMessageBox.warning(self, "Missing file", "Please select REDCap export .csv file")
            return

        if not os.path.isfile(input_path):
            QMessageBox.warning(
                self,
                "Invalid Input",
                "The selected CSV file does not exist.",
            )
            return

        if not input_path.lower().endswith(".csv"):
            QMessageBox.warning(
                self,
                "Invalid Input",
                "The selected input must be a CSV file.",
            )
            return

        if not output_path:
            QMessageBox.warning(self, "Missing output", "Please choose output file location")
            return

        if not output_path.lower().endswith(".xlsx"):
            output_path += ".xlsx"
            self.output_path.setText(output_path)

        output_folder = os.path.dirname(output_path)

        if output_folder and not os.path.isdir(output_folder):
            QMessageBox.warning(
                self,
                "Invalid Output Location",
                "The selected output folder does not exist.",
            )
            return

        if not os.path.isfile(mapping_path):
            QMessageBox.critical(
                self,
                "Missing Configuration",
                f"Mapping file not found:\n{mapping_path}",
            )
            return

        if not os.path.isfile(template_path):
            QMessageBox.critical(
                self,
                "Missing Configuration",
                f"Template file not found:\n{template_path}",
            )
            return

        self.is_running = True

        self.run_btn.setText("Cancel conversion")
        self.run_btn.setEnabled(True)

        self.progress.show()
        self.progress.setValue(0)
        self.progress.setFormat("%p%")

        self.thread = ConvertThread(
            input_path=input_path,
            mapping_path=mapping_path,
            template_path=template_path,
            output_path=output_path,
        )

        self.thread.progress.connect(self.progress.setValue)
        self.thread.success.connect(self.on_finished)
        self.thread.error.connect(self.on_error)
        self.thread.cancelled.connect(self.on_cancelled)

        self.thread.start()

    #Cancel conversion
    def cancel_conversion(self):
        if hasattr(self, "thread"):
            self.thread.stop()

        self.run_btn.setText("Cancelling...")
        self.run_btn.setEnabled(False)

    #Help
    def show_help(self):
        dialog = HelpDialog()
        dialog.exec()

    #Callbacks
    def on_error(self, msg):
        self.is_running = False
        self.run_btn.setText("Run conversion")
        self.run_btn.setEnabled(True)
        self.progress.setValue(0)
        self.progress.setFormat("Failed")
        QMessageBox.information(self, "Error", msg)
        self.progress.hide()

    def on_finished(self):
        self.is_running = False
        self.run_btn.setText("Run conversion")
        self.run_btn.setEnabled(True)
        self.progress.setValue(100)
        self.progress.setFormat("%p%")
        QMessageBox.information(self, "Success!", "Conversion finished!")
        self.progress.hide()

    def on_cancelled(self):
        self.is_running = False
        self.run_btn.setText("Run Conversion")
        self.run_btn.setEnabled(True)
        self.progress.setValue(0)
        self.progress.setFormat("Cancelled")
        QMessageBox.information(self, "Cancelled", "Conversion cancelled.")
        self.progress.hide()