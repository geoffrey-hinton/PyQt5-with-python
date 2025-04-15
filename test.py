import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton
from PyQt5.QtCore import QCoreApplication

class Myapp(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("My first Application")

        btn = QPushButton('Quit', self)
        btn.move(50, 50)
        btn.resize(btn.sizeHint())
        btn.clicked.connect(QCoreApplication.instance().quit)

        self.setWindowTitle("Quit button")

        self.setGeometry(300, 300, 300, 300)
        self.show()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = Myapp()
    sys.exit(app.exec_())
