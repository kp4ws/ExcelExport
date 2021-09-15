import openpyxl
from openpyxl import load_workbook
from datetime import datetime

class ExtractExcelUpdate:

    #Constructor
    def __init__(self):

        #Get user input
        print('Enter absolute path for Workbook "C:\\directory name\\<file name>.xlsx"')
        self._wb_name = input()
        #self._wb = load_workbook('.\\'+self._wb_name+'.xlsx', data_only=True)
        self._wb = load_workbook(self._wb_name, data_only=True)
        self._ws = self._wb.active

        print('Enter absolute path for Export Location "C:\\directory name\\<file name>.sql"')
        self._directory = input()
        
        print('Rows:')
        self._rows = int(input())
        print('Columns:')
        self._columns = int(input())

        #table delimiter
        print('Enter delimiter')
        self._delimiter = input()

    	#Primary method for exporting data
        self.exportData()

    #Export data from excel to sql file
    def exportData(self):
        with open (self._directory, 'w') as f:
            for i in range(2, self._rows+1):
                if(i == 2):
                    f.write('CREATE TABLE table_name (\n')

                data_str = ''
                delim_check = str(self._ws.cell(row=i+1, column=1).value).strip()
                
                for j in range(1, self._columns + 1):
                    if(self._ws.cell(row=i, column=j).value == self._delimiter):
                        break
                    elif(self._ws.cell(row=i,column=j).value is not None):
                        data_str += str(self._ws.cell(row=i, column=j).value).strip() + ' '

                else:
                    if(delim_check == self._delimiter):
                        f.write(data_str + '\n);\n\nCREATE TABLE table_name (\n')
                    else:
                        f.write(data_str[:-1] + ',\n')
            else:
                f.write(');')

                
        print('Complete')


#Main Method
def main():
    ExtractExcelUpdate()

if __name__=='__main__':    
    main()
