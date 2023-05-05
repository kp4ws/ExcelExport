import openpyxl
from openpyxl import load_workbook
from datetime import datetime

class ExtractExcelInsert:
    def __init__(self, wb_name, file, rows, columns, row_start, col_start):
        self._wb_name = wb_name
        self._file = file
        self._rows = rows
        self._columns = columns
        self._row_start = row_start
        self._col_start = col_start

        self._wb = load_workbook(self._wb_name, data_only=True)
        self._ws = self._wb.active
        self.exportData()

    def exportData(self):
        with open (self._file, 'w') as f:
            for i in range(self._row_start, self._rows + self._row_start):
                f.write('INSERT INTO table_name VALUES(')
                data_str = ''

                for j in range(self._col_start, self._columns + self._col_start):
                    
                    if(self._ws.cell(row=i,column=j).value is None or self._ws.cell(row=i,column=j).value == 'NULL'):
                        data_str += 'NULL,'

                    elif (isinstance(self._ws.cell(row=i,column=j).value, int)):
                        data_str += str(self._ws.cell(row=i, column=j).value).strip() + ","

                    else:
                        data_str += "'"+ str(self._ws.cell(row=i, column=j).value).strip() + "',"

                else:
                    f.write(data_str[:-1] + ')\n')