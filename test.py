import sys
from PyQt5.QtWidgets import (
    QApplication, QDialog, QListWidget, QPushButton,
    QVBoxLayout, QLabel, QAbstractItemView
)

class SheetExcludeDialog(QDialog):
    def __init__(self, sheet_names):
        super().__init__()
        self.setWindowTitle("제외할 시트 선택")
        self.setMinimumWidth(300)

        self.all_sheets = sheet_names
        self.remaining_sheets = []

        self.list_widget = QListWidget()
        self.list_widget.addItems(sheet_names)
        self.list_widget.setSelectionMode(QAbstractItemView.MultiSelection)

        self.label = QLabel("제외할 시트를 선택하고 아래 버튼을 누르세요")
        self.confirm_btn = QPushButton("제외 완료")
        self.confirm_btn.clicked.connect(self.exclude_selected)

        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.list_widget)
        layout.addWidget(self.confirm_btn)
        self.setLayout(layout)

    def exclude_selected(self):
        excluded = [item.text() for item in self.list_widget.selectedItems()]
        self.remaining_sheets = [s for s in self.all_sheets if s not in excluded]
        print("남은 시트들:", self.remaining_sheets)
        self.accept()

# 단독 실행 테스트용
if __name__ == "__main__":
    app = QApplication(sys.argv)
    test_sheet_names = [
        "요약", "2023매출", "2023지출", "통계", "그래프", "숨김1"
    ]
    dialog = SheetExcludeDialog(test_sheet_names)
    if dialog.exec_() == QDialog.Accepted:
        print("✔️ 최종 선택된 시트들:", dialog.remaining_sheets)
    sys.exit()