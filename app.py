import json
import sys
from extract_excel_create import ExtractExcelCreate
from extract_excel_insert import ExtractExcelInsert
from extract_excel_update import ExtractExcelUpdate

def main():
    FILE_NAME = '.\config.json'

    def __init__(self):
        self._config = self.retrieveData()
        self._program_type = self._config['program_type']
        self._wb_name = self._config['workbook']
        self._file = self._config['file']
        self._rows = self._config['rows']
        self._columns = self._config['columns']
        self._row_start = self._config['row-start']
        self._col_start = self._config['col-start']
        self._delimiter = self._config['delimiter']
        self.run()

        print('Program Complete')

    def retrieveData(self):
        with open(self.FILE_NAME) as f:
            return json.load(f)

    def run(self):
        if(self._program_type == 'create'):
            ExtractExcelCreate(self._wb_name, self._file, self._rows, self._columns, self._delimiter)
        elif(self._program_type == 'insert'):
            ExtractExcelInsert(self._wb_name, self._file, self._rows, self._columns, self._row_start, self._col_start)
        elif(self._program_type == 'update'):
            ExtractExcelUpdate(self._wb_name, self._file, self._rows, self._columns, self._row_start, self._col_start)
        else: 
              print('***ERROR, invalid program type. Program will now exit.')
              sys.exit()

if __name__=='__main__':    
    main()