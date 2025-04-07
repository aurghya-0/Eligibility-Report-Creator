import sys
from EligibilityReportApp import EligibilityReportApp
from PySide6.QtWidgets import QApplication
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
    # app.setStyleSheet(modern_qss)
    window = EligibilityReportApp()
    window.show()
    sys.exit(app.exec())