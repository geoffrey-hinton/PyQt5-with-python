import sys
import traceback
import pandas as pd
import numpy as np
import resources_rc
import copy
from PyQt5.QtWidgets import *
from PyQt5.QtCore import QAbstractTableModel, Qt
from ui_main_window import Ui_MainWindow
# from functions.page_navigator import *
from functions.separate_location import DistrictGroupApp
from functions.file_handler import read_excel_safely, set_columns, divide_location, divide_spe_location, set_spe_columns, stat_to_excel
from collections import Counter
from itertools import zip_longest

#예외처리 후킹 등록
def except_hook(exctype, value, tb):
    with open("error_log.txt", "w", encoding = 'utf-8') as f:
        f.write("오류 발생\n")
        f.write("".join(traceback.format_exception(exctype, value, tb)))

    msg = QMessageBox()
    msg.setIcon(QMessageBox.Critical)
    msg.setWindowTitle("에러 발생")
    msg.setText("프로그램 실행 중 오류가 발생했습니다.\nerror_log.txt를 확인해주세요")
    msg.exec_()

sys.excepthook = except_hook


# PandasModel 클래스 정의



class PandasModel(QAbstractTableModel):
    def __init__(self, df, parent = None):
        super().__init__(parent)
        self._df = df

    def rowCount(self, parent = None):
        return self._df.shape[0]

    def columnCount(self, parent = None):
        return self._df.shape[1]

    def data(self, index, role = Qt.DisplayRole):
        if index.isValid() and role == Qt.DisplayRole:
            return str(self._df.iloc[index.row(), index.column()])
        return None
    
    def headerData(self, section, orientation, role = Qt.DisplayRole):
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                return str(self._df.columns[section])
            else:
                return str(self._df.index[section])
        return None



class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.stackedWidget.setCurrentIndex(0)

        self.file_path = None
        self.file_handler = None
        self.data = None
        self.group_data = {}
        # self.location = None

        self.page_history = []

        self.scroll_layout = QVBoxLayout(self.scrollAreaWidgetContents)
        self.scrollAreaWidgetContents.setLayout(self.scroll_layout)

        self.scroll_layout_3 = QVBoxLayout(self.scrollAreaWidgetContents_3)
        self.scrollAreaWidgetContents_3.setLayout(self.scroll_layout_3)


        print(self.stackedWidget.currentIndex())
        #버튼 기능 연결
        self.back_btn.clicked.connect(self.go_back(self))
        self.next_btn.clicked.connect(self.go_next(self))
        # self.next_btn.clicked.connect(self.go_to_selected_page)
        self.exit_btn.clicked.connect(self.go_exit(self))

        # 서울, 경기, 부산

        self.seoul_btn.clicked.connect(lambda: self.open_location_window("서울"))
        self.gyeong_btn.clicked.connect(lambda: self.open_location_window("경기"))
        self.busan_btn.clicked.connect(lambda: self.open_location_window("부산"))
        self.btn_browse_3.clicked.connect(self.browse_specific)
        self.start_div_btn.clicked.connect(self.start_div)


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
        self.lineEdit_path_3.textChanged.connect(self.select_column)

        # 지역 및 주식수 확인
        self.btn_browse_4.clicked.connect(self.check_browse)
        self.whole_sheet_3.currentIndexChanged.connect(self.check_num)
        self.enter_stat.clicked.connect(self.stat_file)

        


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
            QApplication.processEvents()
            self.textEdit_3.setPlainText("선택하신 사항에 따라 분할이 진행됩니다.")
            if self.whole_sheet_text in ["서울", "경기", "부산"]:
                whole_df = pd.read_excel(self.lineEdit_path_2.text(), self.whole_sheet_text, dtype = str, skiprows = 2)
            else:
                whole_df = pd.read_excel(self.lineEdit_path_2.text(), self.whole_sheet_text, dtype = str)
            sheets_dict = pd.read_excel(self.lineEdit_path_2.text(), self.selected_sheets, skiprows = 2, dtype = str)
            


            sheets_df = pd.concat(sheets_dict.values(), axis = 0, ignore_index = True)
            col_li = whole_df.columns.to_list()

            merged_df = whole_df.merge(sheets_df, how = "outer", on = col_li, indicator = True)
            
            # 중복 처리
            duplicates = sheets_df[sheets_df.duplicated(subset = col_li, keep = False)]
            duplicates = duplicates.loc[:, ~duplicates.columns.str.startswith("Unnamed")]
            if not duplicates.empty:
                pass
            else:
                self.label_26.setText("중복된 데이터 없습니다.")
                    
            model_dup = PandasModel(duplicates)
            self.tableView_2.setModel(model_dup)
            
            

            # 누락 처리 분기
            omission = merged_df[merged_df["_merge"] == "left_only"].drop(columns = "_merge")
            omission = omission.loc[:, ~omission.columns.str.startswith("Unnamed")]

            if not omission.empty:
                pass
            else:
                self.label_27.setText("누락된 데이터 없습니다.")
            model_omission = PandasModel(omission)
            self.tableView.setModel(model_omission)

    # 분할

    def div_location(self):
        divide_location(self.selected_columns, self.location_text)
        self.ask_return_to_menu()
        self.next_btn.setEnabled(True)
        
        
    # 상세 분할
    def open_location_window(self, location):
        self.location = location
        self.next_btn.setEnabled(True)
        self.start_div_btn.setEnabled(False)
        if location == "서울":
            self.stackedWidget.setCurrentIndex(6)
        elif location == "경기":
            self.stackedWidget.setCurrentIndex(5)
        elif location == "부산":
            self.stackedWidget.setCurrentIndex(7)
        dialog = DistrictGroupApp(location)
        if dialog.exec_() == QDialog.Accepted:
            result = dialog.group_result
            self.selected_columns = copy.deepcopy(result)
            self.group_data[location] = result
            self.temp_dict = result
            for k, v in self.group_data.items():
                print(f"key : {k}")
                print(f"value : {v}")
            print(f"{location} 그룹화 완료:", self.group_data[location])
            dialog.close()
            max_len = max(len(v) for v in self.temp_dict.values())
            for key in self.temp_dict:
                while len(self.temp_dict[key]) < max_len:
                    self.temp_dict[key].append(np.nan)
                    
            temp_df = pd.DataFrame(self.temp_dict)
            temp_df = temp_df.fillna('')
            print(temp_df)
            print(temp_df.shape)
            temp_df_t = temp_df.T
            print(temp_df_t.shape)
            model_temp = PandasModel(temp_df_t)
            self.tableView_3.setModel(model_temp)
            self.tableView_3.horizontalHeader().setStretchLastSection(True)
            self.tableView_3.resizeColumnsToContents()
            self.stackedWidget.setCurrentIndex(8)
    
    # 상세분할 파일 선택 부분
    def browse_specific(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "파일 선택")
        self.lineEdit_path_3.setText(file_path)

        try:
            sheet_names = read_excel_safely(file_path)
            self.next_btn.setEnabled(True)
            self.file_path = file_path
            print(self.file_path)
            self.start_div_btn.setEnabled(True)
        except ValueError as ve:
            QMessageBox.warning(self, "파일 오류", str(ve), QMessageBox.StandardButton.Ok)
            self.next_btn.setEnabled(False)
        except RuntimeError as re:
            QMessageBox.critical(self, "엑셀 읽기 오류", str(re), QMessageBox.StandardButton.Ok)
            self.next_btn.setEnabled(False)

    # 상세분할 열 선택
    def select_column(self):
        self.file_path = self.lineEdit_path_3.text()
        sheet_name = self.location

        try:
            columns = set_spe_columns(self.file_path, sheet_name)
            self.address_combo_4.clear()
            self.sharecount_combo_4.clear()

            self.address_combo_4.addItems(columns)
            self.sharecount_combo_4.addItems(columns)
            self.label_31.setText("주식수, 주소에 해당하는 열과 회사명을 입력해주세요")
        
        except Exception as e:
            QMessageBox.warning(self, "Error", f"{sheet_name} 시트가 없습니다.")
    
    # 상세분할 시작 버튼 부분
    def start_div(self):
        print(self.group_data[self.location])
        selected_columns = {
            "file_path" : self.file_path,
            "sheet_name" : self.location,
            "com_name" : self.lineEdit_2.text(),
            "share_num" : self.sharecount_combo_4.currentText(),
            "address" : self.address_combo_4.currentText(),
            "target_dict" : self.selected_columns
        }
        self.selected_columns = selected_columns

        print(self.selected_columns)


        if "" in selected_columns.values():
            QMessageBox.warning(self, "입력 오류", "모든 항목을 선택(입력) 해주세요")
        else:
            divide_spe_location(self.selected_columns, self.location_text_2)

    def check_browse(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "파일 선택")
        self.file_path = file_path
        self.lineEdit_path_4.setText(file_path)

        try:
            # 이건 직접 만든 함수! 에러는 내부에서 raise 함
            sheet_names = read_excel_safely(file_path)
            self.sheet_names = sheet_names
            print("시트 목록:", self.sheet_names)
            self.whole_sheet_3.clear()
            self.whole_sheet_3.addItems(sheet_names)
            
            self.label_70.setText("전체 시트를 선택해주세요")

        except ValueError as ve:
            QMessageBox.warning(self, "파일 오류", str(ve), QMessageBox.StandardButton.Ok)
            self.next_btn.setEnabled(False)

        except RuntimeError as re:
            QMessageBox.critical(self, "엑셀 읽기 오류", str(re), QMessageBox.StandardButton.Ok)
            self.next_btn.setEnabled(False)

    def check_num(self):
        self.enter_stat.setEnabled(False)
        sheet_names = copy.deepcopy(self.sheet_names)
        while self.scroll_layout_3.count():
            child = self.scroll_layout_3.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
        sheet_names.remove(self.whole_sheet_3.currentText())

        self.sheet_checkboxes = []
        for name in sheet_names:
            checkbox = QCheckBox(name, self.scrollAreaWidgetContents_3)
            self.scroll_layout_3.addWidget(checkbox)
            self.sheet_checkboxes.append(checkbox)
        try:
            sheet_name = self.whole_sheet_3.currentText()
            columns = set_columns(self.file_path, sheet_name)
            self.share_combo_3.clear()
            self.share_combo_3.addItems(columns)
            self.enter_stat.setEnabled(True)

        except Exception as e:
            QMessageBox.warning(self, "Error", f"{sheet_name} 시트의 열을 불러올 수 없습니다.")
            self.enter_stat.setEnabled(False)
        self.label_70.setText("통계 입력을 누르면 입력이 시작됩니다.")
        


    
    def stat_file(self):
        self.whole_sheet_text = self.whole_sheet.currentText()
        self.selected_sheets = [cb.text() for cb in self.sheet_checkboxes if not cb.isChecked()]
        selected_stat = {
            "file_path" : self.file_path,
            "skip_sheets" : self.selected_sheets,
            "whole_sheet" : self.whole_sheet_3.currentText(),
            "share_num" : self.share_combo_3.currentText(),
        }

        if "" in selected_stat.values():
            QMessageBox.warning(self, "입력 오류", "모든 항목을 선택해주세요.")
            return
        else:
            self.selected_stat = selected_stat
            stat_to_excel(self.selected_stat, self.label_70)




    # 기능종료 후 return
    def ask_return_to_menu(self):
        reply = QMessageBox.question(
            self,
            "확인",
            "작업이 끝났습니다 메뉴로 돌아가시겠습니까?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes
        )

        if reply == QMessageBox.Yes:
            self.page_history.append(self.stackedWidget.currentIndex())
            self.stackedWidget.setCurrentIndex(1)

        else:
            pass

    

    # 페이지 이동 관련 함수

    def go_next(self):
        current_index = self.stackedWidget.currentIndex()
        if current_index == 1:
            current_index = self.stackedWidget.currentIndex()
            if self.divide_radio_btn.isChecked():
                self.next_btn.setEnabled(True)
                self.stackedWidget.setCurrentIndex(2)
# 
            elif self.specific_radio_btn.isChecked():
                self.stackedWidget.setCurrentIndex(4)
#   
            elif self.merge_radio_btn.isChecked():
                self.stackedWidget.setCurrentIndex(9)
#   
            elif self.check_omission_btn.isChecked():
                self.stackedWidget.setCurrentIndex(11)
        else:

            print(f"current_index{current_index}")
            self.stackedWidget.setCurrentIndex(current_index + 1)
            self.back_btn.setEnabled(True)
            print(f"current_index : {self.stackedWidget.currentIndex()}")



    def go_back(self):
        print("back clicked")

        current_index = self.stackedWidget.currentIndex()
        self.stackedWidget.setCurrentIndex(current_index - 1)

    def go_exit(self):
        print("Quit")
        QApplication.quit()


        


        

        
    
    
    
    # 파일 이동 관련 함수
    # def go_to_selected_page(self):
        # current_index = self.stackedWidget.currentIndex()
        # if self.divide_radio_btn.isChecked():
            # self.next_btn.setEnabled(True)
            # self.stackedWidget.setCurrentIndex(2)

        # elif self.specific_radio_btn.isChecked():
            # self.stackedWidget.setCurrentIndex(4)

        # elif self.merge_radio_btn.isChecked():
            # self.stackedWidget.setCurrentIndex(6)

        # elif self.check_omission_btn.isChecked():
            # self.stackedWidget.setCurrentIndex(7)

        # elif self.check_num_btn.isChecked():
            # self.stackedWidget.setCurrentIndex(8)
        
def main():
    
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()