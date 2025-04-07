import sys
from EligibilityReportApp import EligibilityReportApp
from PySide6.QtWidgets import QApplication

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = EligibilityReportApp()
    window.show()
    sys.exit(app.exec())