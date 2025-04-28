import sys
import pandas as pd
from PyQt5.QtWidgets import *
from ui_main_window import Ui_MainWindow
import resources_rc
from functions.page_navigator import *
from functions.separate_location import DistrictGroupApp
from functions.file_handler import read_excel_safely, set_columns, divide_location
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

        self.scroll_layout = QVBoxLayout(self.scrollAreaWidgetContents)
        self.scrollAreaWidgetContents.setLayout(self.scroll_layout)


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
        self.save_btn.clicked.connect(self.on_save_clicked)
        self.location_btn.setEnabled(False)

        self.back_btn.setEnabled(False)

        # 분할
        self.location_btn.clicked.connect(self.div_location)

        # 중복 및 누락 확인

        self.btn_browse_2.clicked.connect(self.browse_file_dup)
        self.selected_ok.clicked.connect(self.update_whole_sheet)

    def open_location_window(self, location):
        self.next_btn.setEnabled(False)
        dialog = DistrictGroupApp(location)
        if dialog.exec_() == QDialog.Accepted:
            self.group_data[location] = dialog.group_result
            print(f"{location} 그룹화 완료:", dialog.group_result)
            dialog.close()
        
        

    # 지역 분할 부분 파일 탐색
    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "파일 선택")
        self.lineEdit_path.setText(file_path)

        try:
            # 이건 직접 만든 함수! 에러는 내부에서 raise 함
            sheet_names = read_excel_safely(file_path)
            self.sheet_names = sheet_names
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
            self.address_combo.clear()
            self.sharecount_combo.clear()

            self.address_combo.addItems(columns)
            self.sharecount_combo.addItems(columns)
        
        except Exception as e:
            QMessageBox.warning(self, "Error", f"{sheet_name} 시트 열을 불러올 수 없습니다.")

    def on_save_clicked(self):
        self.next_btn.setEnabled(False)
        selected_columns = {
            "file_path" : self.file_path,
            "sheet_name" : self.sheets_combo.currentText(),
            "com_name" : self.lineEdit.text(),
            "share_num" : self.sharecount_combo.currentText(),
            "address" : self.address_combo.currentText(),
        }

        if "" in selected_columns.values():
            self.location_btn.setEnabled(False)
            QMessageBox.warning(self, "입력 오류", "모든 열을 선택해주세요.")
            return
        else:
            self.location_btn.setEnabled(True)
            

        self.selected_columns = selected_columns
        print("selected_columns", self.selected_columns)

    # 중복 및 누락 확인 부분 타일 탐색

    def browse_file_dup(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "파일 선택")
        self.lineEdit_path_2.setText(file_path)

        try:
            # 이건 직접 만든 함수! 에러는 내부에서 raise 함
            sheet_names = read_excel_safely(file_path)
            self.sheet_names = sheet_names
            print("시트 목록:", self.sheet_names)
            self.next_btn.setEnabled(True)
            self.whole_sheet.clear()
            self.whole_sheet.addItems(sheet_names)
            while self.scroll_layout.count():
                child = self.scroll_layout.takeAt(0)
                if child.widget():
                    child.widget().deleteLater()

            # 새로운 체크박스 추가
            self.sheet_checkboxes = []
            for name in self.sheet_names:
                checkbox = QCheckBox(name, self.scrollAreaWidgetContents)  # 여기 widget parent 지정
                self.scroll_layout.addWidget(checkbox)
                self.sheet_checkboxes.append(checkbox)

        except ValueError as ve:
            QMessageBox.warning(self, "파일 오류", str(ve), QMessageBox.StandardButton.Ok)
            self.next_btn.setEnabled(False)

        except RuntimeError as re:
            QMessageBox.critical(self, "엑셀 읽기 오류", str(re), QMessageBox.StandardButton.Ok)
            self.next_btn.setEnabled(False)

    def update_whole_sheet(self):
        self.whole_sheet_text = self.whole_sheet.currentText()
        self.selected_sheets = [cb.text() for cb in self.sheet_checkboxes if not cb.isChecked()]
        print(f"전체 시트 : {self.whole_sheet_text} \n나머지 시트 리스트 {self.selected_sheets}")
        result = QMessageBox.information(
            self,
            "선택 확인",
            f"비교를 위한 시트:\n{self.whole_sheet_text}\n\n비교 대상 시트들:\n{', '.join(self.selected_sheets)}"
        )

        if result == QMessageBox.Ok:
            self.next_btn.setEnabled(True)

        

    # 분할

    def div_location(self):
        divide_location(self.selected_columns, self.location_text)
        
        


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
        
        

app = QApplication(sys.argv)
window = MainWindow()
window.show()
sys.exit(app.exec_())