import sys
import os
import re
import pandas as pd
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QFileDialog,
    QMessageBox, QLabel, QCheckBox, QLineEdit, QScrollArea
)
from PySide6.QtCore import Qt
from eligibility_processor import extract_subject_codes, process_file

class EligibilityReportApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Eligibility Report Generator")
        self.setGeometry(100, 100, 900, 600)

        self.input_filepath = ""
        self.output_folder_path = ""
        self.subjects = []
        self.selected_subjects = set()

        self.initUI()

    def initUI(self):
        main_layout = QHBoxLayout(self)

        left_panel = QVBoxLayout()

        self.file_label = QLabel("No Excel file selected.")
        left_panel.addWidget(self.file_label)

        self.select_file_btn = QPushButton("Select Excel File")
        self.select_file_btn.clicked.connect(self.select_file)
        left_panel.addWidget(self.select_file_btn)

        self.folder_label = QLabel("No output folder selected.")
        left_panel.addWidget(self.folder_label)

        self.select_folder_btn = QPushButton("Select Output Folder")
        self.select_folder_btn.clicked.connect(self.select_folder)
        left_panel.addWidget(self.select_folder_btn)

        self.combine_checkbox = QCheckBox("Combine Subjects into One PDF")
        left_panel.addWidget(self.combine_checkbox)

        self.export_button = QPushButton("Generate Reports")
        self.export_button.clicked.connect(self.export_reports)
        left_panel.addWidget(self.export_button)

        left_panel.addStretch()

        right_panel = QVBoxLayout()

        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Search subjects...")
        self.search_box.textChanged.connect(self.filter_subjects)
        right_panel.addWidget(self.search_box)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        self.subject_container = QWidget()
        self.subject_layout = QVBoxLayout(self.subject_container)
        self.subject_layout.setAlignment(Qt.AlignTop)
        scroll.setWidget(self.subject_container)

        right_panel.addWidget(scroll)

        main_layout.addLayout(left_panel, 2)
        main_layout.addLayout(right_panel, 3)

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
            for i in reversed(range(self.subject_layout.count())):
                widget = self.subject_layout.itemAt(i).widget()
                if widget is not None:
                    widget.setParent(None)
            for code, name in self.subjects:
                checkbox = QCheckBox(f"{code} - {name}")
                checkbox.setObjectName(code)
                self.subject_layout.addWidget(checkbox)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load subjects: {e}")

    def filter_subjects(self, text):
        for i in range(self.subject_layout.count()):
            widget = self.subject_layout.itemAt(i).widget()
            if isinstance(widget, QCheckBox):
                widget.setVisible(text.lower() in widget.text().lower())

    def export_reports(self):
        selected_codes = []
        for i in range(self.subject_layout.count()):
            widget = self.subject_layout.itemAt(i).widget()
            if isinstance(widget, QCheckBox) and widget.isChecked():
                selected_codes.append(widget.objectName())

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

modern_qss = """
QWidget {
    background-color: #f0f2f5;
    font-family: 'Segoe UI', sans-serif;
    font-size: 12pt;
    color: #2f3640;
}

QPushButton {
    background-color: #1e90ff;
    color: white;
    padding: 10px 18px;
    border-radius: 8px;
    font-weight: bold;
}

QPushButton:hover {
    background-color: #0d78d1;
}

QPushButton:pressed {
    background-color: #095c9d;
}

QCheckBox {
    padding: 6px;
    font-size: 11pt;
}

QCheckBox::indicator {
    width: 18px;
    height: 18px;
}

QCheckBox::indicator:checked {
    background-color: #2ecc71;
    border: 1px solid #27ae60;
    border-radius: 4px;
}

QLineEdit {
    background-color: #ffffff;
    border: 1px solid #ccc;
    padding: 8px;
    border-radius: 6px;
    font-size: 11pt;
}

QLabel {
    font-size: 11pt;
    font-weight: 500;
    padding: 4px 0;
}

QScrollArea {
    background: transparent;
    border: none;
}
"""

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyleSheet(modern_qss)
    window = EligibilityReportApp()
    window.show()
    sys.exit(app.exec())