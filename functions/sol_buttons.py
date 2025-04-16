import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import QCoreApplication

class Myapp(QWidget):

    def __init__(self):
        super().__init__()
        self.button_count = 0
        self.dynamic_buttons = []
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Dynamic Button Creator")
        self.setGeometry(300, 300, 300, 300)

        # Button Layout
        self.button_area_layout = QVBoxLayout()

        # fixed Buttons
        self.create_button = QPushButton("Create Button")
        self.delete_button = QPushButton("Delete Button")

        self.create_button.clicked.connect(self.add_new_button)
        self.delete_button.clicked.connect(self.delete_new_button)

        # create, delete button
        self.bottom_button_layout = QHBoxLayout()
        self.bottom_button_layout.addWidget(self.create_button)
        self.bottom_button_layout.addWidget(self.delete_button)

        # whole layout

        self.main_layout = QVBoxLayout()
        self.main_layout.addLayout(self.button_area_layout)
        self.main_layout.addLayout(self.bottom_button_layout)
        

        self.setLayout(self.main_layout)
        self.show()

    def add_new_button(self):
        if len(self.dynamic_buttons) < 1:
            self.button_count = 0
        self.button_count += 1
        new_button = QPushButton(f"Button {self.button_count}")
        self.button_area_layout.addWidget(new_button)
        self.dynamic_buttons.append(new_button)

    def delete_new_button(self):
        if self.dynamic_buttons:
            btn_to_remove = self.dynamic_buttons.pop()
            self.button_area_layout.removeWidget(btn_to_remove)
            btn_to_remove.deleteLater()
            self.button_count -= 1
            self.update()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = Myapp()
    sys.exit(app.exec_())