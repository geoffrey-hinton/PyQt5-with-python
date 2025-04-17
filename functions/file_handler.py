from PyQt5.QtWidgets import *
import pandas as pd


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