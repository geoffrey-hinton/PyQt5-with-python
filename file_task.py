import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QListWidget, QPushButton,
    QVBoxLayout, QHBoxLayout, QLabel, QScrollArea, QMessageBox
)

class DistrictGrouper(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("서울 구 그룹 나누기")
        self.setGeometry(100, 100, 800, 600)

        self.all_districts = [
            "종로구", "중구", "용산구", "성동구", "광진구", "동대문구", "중랑구", "성북구",
            "강북구", "도봉구", "노원구", "은평구", "서대문구", "마포구", "양천구", "강서구",
            "구로구", "금천구", "영등포구", "동작구", "관악구", "서초구", "강남구", "송파구", "강동구"
        ]

        self.group_count = 0
        self.group_area_widgets = {}

        self.setup_ui()

    def setup_ui(self):
        main_layout = QVBoxLayout()

        # 전체 구 리스트
        self.all_district_list = QListWidget()
        self.all_district_list.addItems(self.all_districts)
        main_layout.addWidget(QLabel("전체 구 목록"))
        main_layout.addWidget(self.all_district_list)

        # 그룹 추가 버튼
        self.add_group_button = QPushButton("그룹 추가")
        self.add_group_button.clicked.connect(self.add_group)
        main_layout.addWidget(self.add_group_button)

        # 그룹 표시 영역 (스크롤 가능)
        self.group_area = QVBoxLayout()
        scroll_area_widget = QWidget()
        scroll_area_widget.setLayout(self.group_area)
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(scroll_area_widget)
        main_layout.addWidget(scroll_area)

        # 아래쪽 버튼들
        bottom_layout = QHBoxLayout()
        self.move_button = QPushButton("→ 그룹으로 이동")
        self.remove_button = QPushButton("← 그룹에서 제거")
        self.complete_button = QPushButton("완료")
        self.complete_button.setEnabled(False)

        self.move_button.clicked.connect(self.move_to_group)
        self.remove_button.clicked.connect(self.move_selected_gu_to_main_list)
        self.complete_button.clicked.connect(self.on_complete)

        bottom_layout.addWidget(self.move_button)
        bottom_layout.addWidget(self.remove_button)
        bottom_layout.addWidget(self.complete_button)

        main_layout.addLayout(bottom_layout)

        self.setLayout(main_layout)

    def add_group(self):
        self.group_count += 1
        group_name = f"서울{self.group_count}"
        group_label = QLabel(group_name)
        group_list = QListWidget()

        self.group_area_widgets[group_name] = group_list

        container = QVBoxLayout()
        container.addWidget(group_label)
        container.addWidget(group_list)
        self.group_area.addLayout(container)

    def move_to_group(self):
        selected_items = self.all_district_list.selectedItems()
        if not selected_items:
            return

        # 가장 최근에 추가된 그룹으로 이동
        if not self.group_area_widgets:
            return

        last_group = list(self.group_area_widgets.keys())[-1]
        group_widget = self.group_area_widgets[last_group]

        for item in selected_items:
            gu = item.text()
            group_widget.addItem(gu)
            self.all_district_list.takeItem(self.all_district_list.row(item))

        self.check_if_all_districts_assigned()

    def move_selected_gu_to_main_list(self):
        for group_name, widget in self.group_area_widgets.items():
            selected_items = widget.selectedItems()
            for item in selected_items:
                gu = item.text()
                # 전체 리스트에 다시 추가
                self.all_district_list.addItem(gu)
                # 그룹에서 제거
                widget.takeItem(widget.row(item))

        self.check_if_all_districts_assigned()

    def check_if_all_districts_assigned(self):
        total_assigned = sum(widget.count() for widget in self.group_area_widgets.values())
        self.complete_button.setEnabled(total_assigned == len(self.all_districts))

    def on_complete(self):
        result = {}
        for group_name, widget in self.group_area_widgets.items():
            result[group_name] = [widget.item(i).text() for i in range(widget.count())]

        msg = "\n".join([f"{k}: {v}" for k, v in result.items()])
        QMessageBox.information(self, "그룹 나누기 완료", msg)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DistrictGrouper()
    window.show()
    sys.exit(app.exec_())