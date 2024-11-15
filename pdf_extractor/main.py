import sys
from PyQt6.QtWidgets import QApplication
from gui import PDFDataExtractor

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFDataExtractor()
    window.show()
    sys.exit(app.exec())
