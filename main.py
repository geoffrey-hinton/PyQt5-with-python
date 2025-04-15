import sys
import pandas as pd
from fuctions.handler import File_handler as Fh
from PyQt5.QtWidgets import *
from ui_main_window import Ui_MainWindow
import resources_rc


class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.stackedWidget.setCurrentIndex(0)

        self.file_path = None
        self.file_handler = None
        self.data = None

        self.page_history = []


        print(self.stackedWidget.currentIndex())
        #버튼 기능 연결
        self.back_btn.clicked.connect(self.go_back)
        self.next_btn.clicked.connect(self.go_to_next_page)
        self.exit_btn.clicked.connect(self.go_exit)

        self.btn_browse.clicked.connect(self.browse_file)
        self.next_btn.clicked.connect(self.go_to_selected_page)


        self.back_btn.setEnabled(False)

        self.find_buttons_in_page()

    
    def find_buttons_in_page(self):
        buttons = self.page_8.findChildren(QPushButton)

        for btn in buttons:
            print(btn.text())
        


    # 함수 모음
    def go_to_next_page(self):
        print("clicked_next")
        current_index = self.stackedWidget.currentIndex()
        max_index = self.stackedWidget.count() - 1

        if current_index < max_index:
            self.stackedWidget.setCurrentIndex(current_index + 1)
        current_index = self.stackedWidget.currentIndex()

        self.next_btn.setEnabled(current_index < max_index)

        self.back_btn.setEnabled(current_index > 0)

    def go_back(self):
        print("clicked_back")
        if self.page_history:
            last_page = self.page_history.pop()
            self.stackedWidget.setCurrentIndex(last_page)

        self.back_btn.setEnabled(len(self.page_history) > 0)
        

    def go_exit(self):
        QApplication.quit()

    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "파일 선택")
        
        if file_path:
            self.lineEdit_path.setText(file_path)
            if file_path[-4:] == "xlsx":
                self.file_path = file_path
                self.file_handler = Fh(self.file_path)
                try:
                    self.file_handler.read_file()
                    self.data = self.file_handler.get_data()
                    self.next_btn.setEnabled(True)
                except PermissionError as e:
                    QMessageBox.warning(
                        self, "파일 오류",
                        "파일이 열려있습니다 닫고 다시 시도해주세요", QMessageBox.StandardButton.Ok
                    )
                    self.next_btn.setEnabled(False)
                except FileNotFoundError as e:
                    QMessageBox.warning(
                        self, "파일 오류",
                        "파일을 찾을 수 없습니다 다시 시도해주세요", QMessageBox.StandardButton.Ok
                    )
                    self.next_btn.setEnabled(False)
                except Exception as e:
                    QMessageBox.warning(
                        self, "파일 오류",
                        f"예상치 못한 오류가 발생했습니다 다시 시도해주세요\n{e}",
                        QMessageBox.StandardButton.Ok
                    )
                    self.next_btn.setEnabled(False)
            else:
                QMessageBox.warning(
                    self, "파일 오류", 
                    "엑셀 파일이 아닙니다 확장자를 확인해주세요", QMessageBox.StandardButton.Ok
                    )
                self.next_btn.setEnabled(False)
    def go_to_selected_page(self):
        current_index = self.stackedWidget.currentIndex()
        self.page_history.append(current_index)
        if self.divide_radio_btn.isChecked():
            self.stackedWidget.setCurrentIndex(2)

        elif self.specific_radio_btn.isChecked():
            self.stackedWidget.setCurrentIndex(4)

        elif self.merge_radio_btn.isChecked():
            self.stackedWidget.setCurrentIndex(6)

        elif self.check_omission_btn.isChecked():
            self.stackedWidget.setCurrentIndex(7)

        elif self.check_num_btn.isChecked():
            self.stackedWidget.setCurrentIndex(8)
        
    def go_to_specific_split(self):
        pass
        

app = QApplication(sys.argv)
window = MainWindow()
window.show()
sys.exit(app.exec_())