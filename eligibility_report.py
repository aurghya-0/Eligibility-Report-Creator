import sys
import os
import re
import pandas as pd
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog,
    QMessageBox, QLabel, QListWidget, QListWidgetItem, QCheckBox, QLineEdit
)
from PySide6.QtCore import Qt
from eligibility_processor import extract_subject_codes, process_file

class EligibilityReportApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Eligibility Report Generator")
        self.setGeometry(100, 100, 600, 500)

        self.input_filepath = ""
        self.output_folder_path = ""
        self.subjects = []
        self.selected_subjects = set()

        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.file_label = QLabel("No Excel file selected.")
        layout.addWidget(self.file_label)

        self.select_file_btn = QPushButton("Select Excel File")
        self.select_file_btn.clicked.connect(self.select_file)
        layout.addWidget(self.select_file_btn)

        self.folder_label = QLabel("No output folder selected.")
        layout.addWidget(self.folder_label)

        self.select_folder_btn = QPushButton("Select Output Folder")
        self.select_folder_btn.clicked.connect(self.select_folder)
        layout.addWidget(self.select_folder_btn)

        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Search subjects...")
        self.search_box.textChanged.connect(self.filter_subjects)
        layout.addWidget(self.search_box)

        self.subject_list = QListWidget()
        layout.addWidget(self.subject_list)

        self.combine_checkbox = QCheckBox("Combine Subjects into One PDF")
        layout.addWidget(self.combine_checkbox)

        self.export_button = QPushButton("Generate Reports")
        self.export_button.clicked.connect(self.export_reports)
        layout.addWidget(self.export_button)

        self.setLayout(layout)

    def select_file(self):
        filepath, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")
        if filepath:
            self.input_filepath = filepath
            self.file_label.setText(f"Selected File: {os.path.basename(filepath)}")
            self.load_subjects()

    def select_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder_path:
            self.output_folder_path = folder_path
            self.folder_label.setText(f"Selected Output Folder: {folder_path}")

    def load_subjects(self):
        try:
            self.subjects, self.df = extract_subject_codes(self.input_filepath)
            self.subject_list.clear()
            for code, name in self.subjects:
                item = QListWidgetItem(f"{code} - {name}")
                item.setData(Qt.UserRole, code)
                item.setCheckState(Qt.Unchecked)
                self.subject_list.addItem(item)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load subjects: {e}")

    def filter_subjects(self, text):
        for i in range(self.subject_list.count()):
            item = self.subject_list.item(i)
            item.setHidden(text.lower() not in item.text().lower())

    def export_reports(self):
        selected_codes = [self.subject_list.item(i).data(Qt.UserRole)
                          for i in range(self.subject_list.count())
                          if self.subject_list.item(i).checkState() == Qt.Checked]

        if not self.input_filepath or not self.output_folder_path:
            QMessageBox.warning(self, "Warning", "Please select both input file and output folder.")
            return

        if not selected_codes:
            QMessageBox.warning(self, "Warning", "Please select at least one subject.")
            return

        try:
            process_file(self.df, selected_codes, self.output_folder_path, self.combine_checkbox.isChecked())
            QMessageBox.information(self, "Success", "Reports generated successfully!")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to generate reports: {e}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = EligibilityReportApp()
    window.show()
    sys.exit(app.exec())
