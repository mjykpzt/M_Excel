from data_random.dataConfig import tem_random
from e_io import excel_io


class TemperatureXlsx(excel_io.WorkBook):
    def __init__(self, path,sheet_name, row, col,dataC=tem_random.Build().build()):
        super().__init__(path)
        super().open_sheet(sheet_name)
        self.base = None
        self.data = dataC
        self.row = row
        self.col = col
        self.sheet_name = sheet_name
        self.num = None

    def exe(self):
        row = self.row
        col = self.col
        for i in range(len(self.base)):
            super().write_by_index(row, col + i, self.base[i])
        col += len(self.base)
        for i in range(self.num):
            super().write_by_index(row, col + i, str(self.data.random()) + "℃")
        self.row += 1

    def add(self,base,num,dataC=tem_random.Build().build()):
        self.base=base
        self.num=num
        self.data = dataC

    def set_start(self,row,col):
        self.row = row
        self.col = col


