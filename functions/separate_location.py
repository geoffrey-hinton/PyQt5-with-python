import sys
from PyQt5.QtWidgets import *

class DistrictGroupApp(QDialog):
    def __init__(self, location):
        super().__init__()
        self.location = location
        self.setWindowTitle(f"{location}분류")
        self.setGeometry(100, 100, 700, 450)
        self.group_counter = 2  # 서울1, 서울2는 기본 그룹
        if location == "서울":
            self.all_districts = sorted([
                "강남구", "강동구", "강북구", "강서구", "관악구", "광진구", "구로구", "금천구",
                "노원구", "도봉구", "동대문구", "동작구", "마포구", "서대문구", "서초구", "성동구",
                "성북구", "송파구", "양천구", "영등포구", "용산구", "은평구", "종로구", "중구", "중랑구"
            ])
        elif location == "부산":
            self.all_districts = sorted([
                "중구", "서구", "동구", "영도구", "부산진구", "동래구", "남구", "북구",
                "해운대구", "사하구", "금정구", "강서구", "연제구", "수영구", "사상구", "기장군"
            ])
        
        elif location == "경기":
            self.all_districts = sorted([
                "수원시", "성남시", "의정부시", "안양시", "부천시", "광명시", "평택시",
                "동두천시", "안산시", "고양시", "과천시", "구리시", "남양주시", "오산시",
                "시흥시", "군포시", "의왕시", "하남시", "용인시", "파주시", "이천시", "안성시",
                "김포시", "화성시", "광주시", "양주시", "포천시", "여주시", "연천군", "가평군", "양평군"
            ])
            

        self.all_district_list = QListWidget()
        self.all_district_list.addItems(self.all_districts)
        self.all_district_list.setSortingEnabled(True)

        self.group_combobox = QComboBox()
        self.group_combobox.addItems([f"{location}1", f"{location}2"])
        self.group_combobox.currentTextChanged.connect(self.update_group_display)

        self.group_widgets = {
            f"{location}1": QListWidget(),
            f"{location}2": QListWidget()
        }

        self.group_display = self.group_widgets[f"{location}1"]

        self.add_button = QPushButton("구 선택 → 그룹 추가")
        self.remove_button = QPushButton("그룹 → 구 리스트")
        self.finish_button = QPushButton("완료")
        self.add_group_button = QPushButton("그룹 추가")
        self.delete_group_button = QPushButton("그룹 삭제")

        self.finish_button.setEnabled(False)

        self.init_ui()

    def init_ui(self):
    # 왼쪽 레이아웃
        left_layout = QVBoxLayout()
        left_layout.addWidget(QLabel("전체 구 리스트"))
        left_layout.addWidget(self.all_district_list)
    
        # 가운데 레이아웃
        middle_layout = QVBoxLayout()
        middle_layout.addWidget(self.add_button)
        middle_layout.addWidget(self.remove_button)
        middle_layout.addWidget(self.finish_button)
    
        # 오른쪽 레이아웃
        right_layout = QVBoxLayout()
        right_layout.addWidget(QLabel("그룹 선택"))
    
        # 그룹 선택 + 추가/삭제 버튼 가로 배치
        group_selection_layout = QHBoxLayout()
        group_selection_layout.addWidget(self.group_combobox)
        group_selection_layout.addWidget(self.add_group_button)
        group_selection_layout.addWidget(self.delete_group_button)
        right_layout.addLayout(group_selection_layout)
    
        right_layout.addWidget(QLabel("선택된 그룹 구 리스트"))
        right_layout.addWidget(self.group_display)
    
        # 메인 전체 레이아웃
        main_layout = QHBoxLayout()
        main_layout.addLayout(left_layout)
        main_layout.addLayout(middle_layout)
        main_layout.addLayout(right_layout)
        self.setLayout(main_layout)
    
        # 버튼 기능 연결
        self.add_button.clicked.connect(self.move_selected_gu_to_group)
        self.remove_button.clicked.connect(self.move_gu_back_to_all)
        self.add_group_button.clicked.connect(self.add_group)
        self.delete_group_button.clicked.connect(self.delete_selected_group)
        self.finish_button.clicked.connect(self.finish_grouping)


    # 지역 나누기
    def update_group_display(self, group_name):
        if hasattr(self, "group_display"):
            self.layout().itemAt(2).layout().replaceWidget(self.group_display, self.group_widgets[group_name])
            self.group_display.hide()
        self.group_display = self.group_widgets[group_name]
        self.group_display.show()

        # 기본 그룹은 삭제 비활성화
        if group_name in ["1", "2"]:
            self.delete_group_button.setEnabled(False)
        else:
            self.delete_group_button.setEnabled(True)

    def move_selected_gu_to_group(self):
        selected_items = self.all_district_list.selectedItems()
        if selected_items:
            gu = selected_items[0].text()
            group_name = self.group_combobox.currentText()
            if gu not in [self.group_widgets[g].item(i).text() for g in self.group_widgets for i in range(self.group_widgets[g].count())]:
                self.group_widgets[group_name].addItem(gu)
                self.all_district_list.takeItem(self.all_district_list.row(selected_items[0]))
                self.check_if_all_districts_assigned()

    def move_gu_back_to_all(self):
        group_name = self.group_combobox.currentText()
        selected_items = self.group_widgets[group_name].selectedItems()
        if selected_items:
            gu = selected_items[0].text()
            self.all_district_list.addItem(gu)
            self.group_widgets[group_name].takeItem(self.group_widgets[group_name].row(selected_items[0]))
            self.all_district_list.sortItems()
            self.check_if_all_districts_assigned()

    def check_if_all_districts_assigned(self):
        total_selected = sum(group.count() for group in self.group_widgets.values())
        if total_selected == len(self.all_districts):
            self.finish_button.setEnabled(True)
        else:
            self.finish_button.setEnabled(False)

    def add_group(self):
        self.group_counter += 1
        group_name = self.location + str(self.group_counter)
        new_list = QListWidget()
        self.group_widgets[group_name] = new_list
        self.group_combobox.addItem(group_name)
        print(self.group_counter)

    def delete_selected_group(self):
        group_name = self.group_combobox.currentText()
        if group_name in [f"{self.location}1", f"{self.location}2"]:
            self.delete_group_button.setEnabled(False)
            return  # 삭제 불가

        # 구 다시 전체 리스트로
        for i in range(self.group_widgets[group_name].count()):
            self.all_district_list.addItem(self.group_widgets[group_name].item(i).text())
        self.all_district_list.sortItems()

        # UI에서 제거
        self.group_widgets[group_name].hide()
        self.group_combobox.removeItem(self.group_combobox.currentIndex())
        del self.group_widgets[group_name]

        self.group_counter -= 1
        print(self.group_counter)

        # 그룹 번호 재정렬은 선택사항 (지금은 이름만 유지)

        self.check_if_all_districts_assigned()
        self.update_group_display(f"{self.location}1")  # 기본으로 돌아감

    def finish_grouping(self):
        result = {}
        for group_name, widget in self.group_widgets.items():
            loc_list = [widget.item(i).text() for i in range(widget.count())]
            result[group_name] = loc_list

        self.group_result = result

        # print("그룹화 결과", self.group_result)
        self.accept()
