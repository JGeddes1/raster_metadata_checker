import os
import sys
import openpyxl
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QHBoxLayout, QFileDialog, QListWidget, QMessageBox, QGridLayout

class FileChecker(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('File Checker')
        self.setGeometry(100, 100, 800, 600)
        
        # Directory selection
        self.directory_label = QLabel('Select Directory:', self)
        self.directory_path = QLineEdit(self)
        self.browse_directory_button = QPushButton('Browse', self)
        self.browse_directory_button.clicked.connect(self.browse_directory)

        # Excel file selection
        self.excel_label = QLabel('Select Excel File:', self)
        self.excel_file_path = QLineEdit(self)
        self.browse_excel_button = QPushButton('Browse', self)
        self.browse_excel_button.clicked.connect(self.browse_excel_file)

        # Second Excel file selection
        self.excel_label2 = QLabel('Select Second Excel File:', self)
        self.excel_file_path2 = QLineEdit(self)
        self.browse_excel_button2 = QPushButton('Browse', self)
        self.browse_excel_button2.clicked.connect(self.browse_excel_file2)

        # Check files button
        self.check_button = QPushButton('Check Files', self)
        self.check_button.clicked.connect(self.check_files)

        # Listboxes for missing files and keywords
        self.missing_directory_label = QLabel('Missing from directory:', self)
        self.missing_directory_listbox = QListWidget(self)
        
        self.missing_metadata_label = QLabel('Not listed in metadata:', self)
        self.missing_metadata_listbox = QListWidget(self)
        
        self.keyword_label = QLabel('Subject Keywords:', self)
        self.keyword_listbox = QListWidget(self)

        # Status label
        self.status_label = QLabel('', self)

        # Layouts
        main_layout = QVBoxLayout()

        directory_layout = QHBoxLayout()
        directory_layout.addWidget(self.directory_label)
        directory_layout.addWidget(self.directory_path)
        directory_layout.addWidget(self.browse_directory_button)
        main_layout.addLayout(directory_layout)

        excel_layout = QHBoxLayout()
        excel_layout.addWidget(self.excel_label)
        excel_layout.addWidget(self.excel_file_path)
        excel_layout.addWidget(self.browse_excel_button)
        main_layout.addLayout(excel_layout)

        excel_layout2 = QHBoxLayout()
        excel_layout2.addWidget(self.excel_label2)
        excel_layout2.addWidget(self.excel_file_path2)
        excel_layout2.addWidget(self.browse_excel_button2)
        main_layout.addLayout(excel_layout2)

        main_layout.addWidget(self.check_button)

        listbox_layout = QGridLayout()
        listbox_layout.addWidget(self.missing_directory_label, 0, 0)
        listbox_layout.addWidget(self.missing_directory_listbox, 1, 0)
        listbox_layout.addWidget(self.missing_metadata_label, 0, 1)
        listbox_layout.addWidget(self.missing_metadata_listbox, 1, 1)
        listbox_layout.addWidget(self.keyword_label, 0, 2)
        listbox_layout.addWidget(self.keyword_listbox, 1, 2)
        main_layout.addLayout(listbox_layout)

        main_layout.addWidget(self.status_label)

        self.setLayout(main_layout)

    def browse_directory(self):
        directory = QFileDialog.getExistingDirectory(self, 'Select Directory')
        if directory:
            self.directory_path.setText(directory)

    def browse_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, 'Select Excel File', '', 'Excel files (*.xlsx *.xls)')
        if file_path:
            self.excel_file_path.setText(file_path)

    def browse_excel_file2(self):
        file_path, _ = QFileDialog.getOpenFileName(self, 'Select Second Excel File', '', 'Excel files (*.xlsx *.xls)')
        if file_path:
            self.excel_file_path2.setText(file_path)

    def read_excel_file(self, file_path):
        filenames = []
        workbook = openpyxl.load_workbook(file_path)
        worksheet = workbook.active
        for row in worksheet.iter_rows(values_only=True):
            if row[0] and not str(row[0]).lower() in ["Filename", "filename"]:
                filenames.append(row[0].strip().lower())
            if row[2] and not str(row[2]).lower() in ["Subject Keyword 1", "subjectword"]:
                subjectkeywords.add(row[2].strip().lower())
            if row[3] and not str(row[3]).lower() in ["Subject Keyword 2", "subjectword"]:
                subjectkeywords.add(row[3].strip().lower())
            if row[4] and not str(row[4]).lower() in ["Subject Keyword 3", "subjectword"]:
                subjectkeywords.add(row[4].strip().lower())
        return filenames

    def list_files(self, directory):
        image_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.tif']
        return [filename.lower() for filename in os.listdir(directory) if os.path.isfile(os.path.join(directory, filename)) and os.path.splitext(filename)[1].lower() in image_extensions]

    def find_missing_files(self, directory_files, excel_filenames):
        missing_in_directory = [filename for filename in excel_filenames if filename not in directory_files]
        missing_in_metadata = [filename for filename in directory_files if filename not in excel_filenames]
        return missing_in_directory, missing_in_metadata

    def check_files(self):
        directory = self.directory_path.text()
        excel_file1 = self.excel_file_path.text()
        excel_file2 = self.excel_file_path2.text()

        if not directory or not excel_file1:
            QMessageBox.warning(self, 'Error', 'Please select a directory and at least one Excel file.')
            return

        global subjectkeywords
        subjectkeywords = set()

        excel_filenames1 = self.read_excel_file(excel_file1)
        excel_filenames2 = self.read_excel_file(excel_file2) if excel_file2 else []
        excel_filenames = list(set(excel_filenames1 + excel_filenames2))

        directory_files = self.list_files(directory)

        missing_in_directory, missing_in_metadata = self.find_missing_files(directory_files, excel_filenames)

        self.missing_directory_listbox.clear()
        self.missing_metadata_listbox.clear()
        self.keyword_listbox.clear()

        keyword_remove_list = ['subject keyword 3', 'subject keyword 2', 'subject keyword 1']

        for file in missing_in_directory:
            self.missing_directory_listbox.addItem(file)
        for file in missing_in_metadata:
            self.missing_metadata_listbox.addItem(file)
        for keyword in subjectkeywords:
            if keyword not in keyword_remove_list:
                self.keyword_listbox.addItem(keyword)

        if not missing_in_directory and not missing_in_metadata:
            self.status_label.setText("No missing files found! All good!")
            self.status_label.setStyleSheet("color: green;")
        else:
            self.status_label.setText("Missing files detected. Check the lists above.")
            self.status_label.setStyleSheet("color: red;")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = FileChecker()
    ex.show()
    sys.exit(app.exec_())
