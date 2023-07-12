import openpyxl


class ExcelControl():

    def __init__(self, excel_name = None):

        self._workbook = openpyxl.Workbook()
        
        if excel_name != None :
            self._workbook = openpyxl.load_workbook(filename=excel_name)


        print(self.workbook.sheetnames)
        print("Active: ", self.workbook.active.title)
        self._worksheet = self.workbook.active
        print(self.worksheet)
        print(self.worksheet.cell(row=1 , column=1).value)
        self._cell = self.worksheet.cell(row=1 , column=1)
        print(self.cell)


    @property
    def workbook(self) -> openpyxl.Workbook:
        return self._workbook
    
    @workbook.setter
    def workbook_setter(self, workbook):
        self._workbook = workbook

    @property
    def worksheet(self):
        return self._worksheet
    
    @worksheet.setter
    def worksheet_setter(self, worksheet):
        self._worksheet = worksheet
    
    @property
    def cell(self):
        return self._cell

if __name__ == "__main__":
    ev = ExcelControl()
    
