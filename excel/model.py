import openpyxl
import os

class ExcelModel:

    def __init__(self):
        self._data = ""

    def readExcel(self, file_path):
        self._data = openpyxl.load_workbook(file_path, data_only=True)

    def read_work_book(self, sheet_name):
        sheet = self._data[sheet_name]
        result = []
        for row in sheet.rows:
            result.append([row[0].value, row[-2].value])
        return result

    def write_csv(self, fail_list):     #fail_list는 엑셀에 저장할 데이터 목록
        wb = openpyxl.workbook.Workbook()
        fail_work = wb.active
        fail_work.title = "fail"
        for fail in fail_list: # 데이터를 엑셀파일에 기록한다
            fail_work.append(fail)
        wb.save(self._file_path)
        wb.close()
