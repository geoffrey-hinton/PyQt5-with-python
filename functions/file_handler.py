import pandas as pd
from PyQt5.QtWidgets import *

def read_excel_safely(file_path):
    """엑셀 파일 경로로부터 시트 목록을 안전하게 반환"""
    if not file_path.lower().endswith((".xlsx", ".xls")):
        raise ValueError("엑셀 파일이 아닙니다. 확장자 확인 요망.")

    try:
        xl = pd.ExcelFile(file_path)
        return xl.sheet_names
    except Exception as e:
        raise RuntimeError(f"엑셀 파일 읽기 오류: {str(e)}")
    

def set_columns(file_path, sheet_name):
    try:
        target_file = pd.read_excel(file_path, sheet_name = sheet_name)
        print(target_file.columns)
        return list(target_file.columns)
    except Exception as e:
        raise e
    

def divide_location(div_dict, label):
    df = pd.read_excel(div_dict["file_path"], sheet_name = div_dict["sheet_name"])
    df.astype(str)
    com_name = div_dict['com_name']
    share_num = div_dict['share_num']
    address = div_dict["address"]
    df[share_num] = df[share_num].astype(int)
    
