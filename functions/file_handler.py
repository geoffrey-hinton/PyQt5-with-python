import pandas as pd

def read_excel_safely(file_path):
    """엑셀 파일 경로로부터 시트 목록을 안전하게 반환"""
    if not file_path.lower().endswith((".xlsx", ".xls")):
        raise ValueError("엑셀 파일이 아닙니다. 확장자 확인 요망.")

    try:
        xl = pd.ExcelFile(file_path)
        return xl.sheet_names
    except Exception as e:
        raise RuntimeError(f"엑셀 파일 읽기 오류: {str(e)}")