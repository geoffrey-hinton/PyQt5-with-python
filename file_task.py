import sys
from PyQt5.QtWidgets import QApplication, QWidget, QListWidget, QVBoxLayout, QHBoxLayout, QPushButton, QComboBox, QLabel

class DistrictGroupApp(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("서울 구 분류")
        self.setGeometry(100, 100, 600, 400)
        self.group_list = []

        self.all_districts = [
            "강남구", "강동구", "강북구", "강서구", "관악구", "광진구", "구로구", "금천구",
            "노원구", "도봉구", "동대문구", "동작구", "마포구", "서대문구", "서초구", "성동구",
            "성북구", "송파구", "양천구", "영등포구", "용산구", "은평구", "종로구", "중구", "중랑구"
        ]

        # UI 요소들
        self.all_district_list = QListWidget()  # 전체 구 리스트
        self.group_area_combobox = QComboBox()  # 그룹 선택 콤보박스
        self.group_area_combobox.addItems(["서울1", "서울2", "서울3", "서울4"])

        self.group_area_widgets = {
            "서울1": QListWidget(),
            "서울2": QListWidget(),
            "서울3": QListWidget(),
            "서울4": QListWidget()
        }

        self.add_button = QPushButton("구 선택 → 그룹 추가")
        self.finish_button = QPushButton("완료")
        self.finish_button.setEnabled(False)

        # 레이아웃 설정
        self.init_ui()

    def init_ui(self):
        # 전체 구 리스트 채우기
        self.all_district_list.addItems(self.all_districts)

        # 레이아웃 설정
        left_layout = QVBoxLayout()
        left_layout.addWidget(QLabel("전체 구 리스트"))
        left_layout.addWidget(self.all_district_list)

        right_layout = QVBoxLayout()
        right_layout.addWidget(QLabel("그룹 선택"))
        right_layout.addWidget(self.group_area_combobox)
        right_layout.addWidget(QLabel("그룹 리스트"))
        for group_name, widget in self.group_area_widgets.items():
            right_layout.addWidget(widget)

        bottom_layout = QHBoxLayout()
        bottom_layout.addWidget(self.add_button)
        bottom_layout.addWidget(self.finish_button)

        main_layout = QHBoxLayout()
        main_layout.addLayout(left_layout)
        main_layout.addLayout(right_layout)
        main_layout.addLayout(bottom_layout)

        self.setLayout(main_layout)

        # 버튼 연결
        self.add_button.clicked.connect(self.move_selected_gu_to_group)
        self.finish_button.clicked.connect(self.finish_grouping)

    def move_selected_gu_to_group(self):
        selected_items = self.all_district_list.selectedItems()
        if selected_items:
            gu = selected_items[0].text()
            group_name = self.group_area_combobox.currentText()

            # 그룹에 추가
            self.group_area_widgets[group_name].addItem(gu)
            # 전체 리스트에서 제거
            self.all_district_list.takeItem(self.all_district_list.row(selected_items[0]))

            self.check_if_all_districts_assigned()

    def check_if_all_districts_assigned(self):
        total_selected = sum(group.count() for group in self.group_area_widgets.values())
        if total_selected == len(self.all_districts):
            self.finish_button.setEnabled(True)
        else:
            self.finish_button.setEnabled(False)

    def finish_grouping(self):
        # 완료 버튼 클릭 시 처리
        print("그룹화 완료!")
        # 여기서 창을 닫거나 다음 단계로 넘어갈 수 있도록 구현 가능
        self.close()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = DistrictGroupApp()
    window.show()
    sys.exit(app.exec_())