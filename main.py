import sys
import pandas as pd
from PyQt5.QtWidgets import *
from ui_main_window import Ui_MainWindow
import resources_rc
from functions.page_navigator import *
from functions.separate_location import DistrictGroupApp
from functions.file_handler import read_excel_safely, set_columns
import pandas as pd


class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.stackedWidget.setCurrentIndex(0)

        self.file_path = None
        self.file_handler = None
        self.data = None
        self.group_data = {}

        self.page_history = []


        print(self.stackedWidget.currentIndex())
        #버튼 기능 연결
        self.back_btn.clicked.connect(lambda: go_back(self))
        self.next_btn.clicked.connect(lambda: go_next(self))
        self.next_btn.clicked.connect(self.go_to_selected_page)
        self.exit_btn.clicked.connect(lambda: go_exit(self))

        # 서울, 경기, 부산

        self.seoul_btn.clicked.connect(lambda: self.open_location_window("서울"))
        self.gyeong_btn.clicked.connect(lambda: self.open_location_window("경기"))
        self.busan_btn.clicked.connect(lambda: self.open_location_window("부산"))


        # 파일
        self.btn_browse.clicked.connect(self.browse_file)
        self.sheets_combo.currentIndexChanged.connect(self.on_sheet_selected)
        self.save_btn.cliked.connect(self.on_save_clicked)

        self.back_btn.setEnabled(False)

    def open_location_window(self, location):
        self.next_btn.setEnabled(False)
        dialog = DistrictGroupApp(location)
        if dialog.exec_() == QDialog.Accepted:
            self.group_data[location] = dialog.group_result
            print(f"{location} 그룹화 완료:", dialog.group_result)
            dialog.close()
        
        


    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "파일 선택")
        self.lineEdit_path.setText(file_path)

        try:
            # 이건 직접 만든 함수! 에러는 내부에서 raise 함
            sheet_names = read_excel_safely(file_path)
            self.file_path = file_path
            self.sheet_names = sheet_names
            print("시트 목록:", self.sheet_names)
            self.next_btn.setEnabled(True)
            self.sheets_combo.clear()
            self.sheets_combo.addItems(sheet_names)

        except ValueError as ve:
            QMessageBox.warning(self, "파일 오류", str(ve), QMessageBox.StandardButton.Ok)
            self.next_btn.setEnabled(False)

        except RuntimeError as re:
            QMessageBox.critical(self, "엑셀 읽기 오류", str(re), QMessageBox.StandardButton.Ok)
            self.next_btn.setEnabled(False)


    # 시트에 포함된 열 불러오기
    def on_sheet_selected(self):
        sheet_name = self.sheets_combo.currentText()

        try:
            columns = set_columns(self.file_path, sheet_name)
            self.zip_code_combo.clear()
            self.address_combo.clear()
            self.sharecount_combo.clear()

            self.zip_code_combo.addItems(columns)
            self.address_combo.addItems(columns)
            self.sharecount_combo.addItems(columns)
        
        except Exception as e:
            QMessageBox.warning(self, "Error", f"{sheet_name} 시트 열을 불러올 수 없습니다.")
        




    def go_to_selected_page(self):
        current_index = self.stackedWidget.currentIndex()
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