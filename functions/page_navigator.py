from PyQt5.QtWidgets import *

def go_next(main_window):
    

    current_index = main_window.stackedWidget.currentIndex()
    print(f"current_index{current_index}")
    main_window.stackedWidget.setCurrentIndex(current_index + 1)
    main_window.back_btn.setEnabled(True)
    print(f"current_index : {main_window.stackedWidget.currentIndex()}")
    


def go_back(main_window):
    print("back clicked")
    
    current_index = main_window.stackedWidget.currentIndex()
    main_window.stackedWidget.setCurrentIndex(current_index - 1)

def go_exit(main_window):
    print("Quit")
    QApplication.quit()