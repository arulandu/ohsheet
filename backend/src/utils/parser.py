import xlwings as xw

class ExcelBook:
    def __init__(self, file_path: str):
        self.wb = xw.Book(file_path)
    
    def get_sheet(self, sheet_index: int):
        return self.wb.sheets[sheet_index - 1]
   
    def get_range_values(self, sheet_index: int, range_address: str):
        sheet = self.get_sheet(sheet_index)
        range_obj = sheet.range(range_address)
        return range_obj.value
    
    def get_sheet_index(self, sheet_name: str):
        return self.wb.sheets.index(sheet_name) + 1
    
    def close(self):
        self.wb.close()
