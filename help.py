from PySide6.QtWidgets import QDialog, QVBoxLayout, QTextEdit

class HelpDialog(QDialog):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("How to use")
        self.resize(800, 600)

        layout = QVBoxLayout()

        self.help_text = QTextEdit()
        self.help_text.setReadOnly(True)

        self.help_text.setPlainText("""
        
        Databank Export Converter
            
        The converter will take in the labels-based .csv file exported from the Physical Exam section of the REDCap Database.
        It is currently only set up to read and convert the label-based file, as opposed to the raw-data-based file.
            
        To export the label-based file:
             1. In REDCap, navigate to the "Data Exports, Reports, and Stats" application.
             2. Under "My Reports & Exports", click the "Export data" button for the Physical Exam.
             3. Select the "CSV/Microsoft Excel (labels)" option and click "Export Data".
             4. Click the "EXCEL CSV Labels" icon with the arrow to download.
        
        How to use the converter:
             1. Click "Load CSV File" to load your label-based REDCap .csv export file. Alternatively, you can drag and drop a file into the box.
             2. Click "Save Output As" to choose where to save your converted file.
             3. Click "Run Converter" to run the converter
             4. Converter should run successfully.
             
        The converter is a work in progress.  
        """)

        layout.addWidget(self.help_text)
        self.setLayout(layout)