from PyQt6.QtWidgets import QApplication
from window import Window

if __name__ == "__main__":
    app = QApplication([])  # Initialize the QApplication instance
    start = Window()       # Create the window instance
    app.exec() 