import openpyxl
from openpyxl import load_workbook

class ExtractExcel:

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

    	#Primary method for exporting data
        self.exportData()

    #Export data from excel to sql file
    def exportData(self):
        with open (self._directory, 'w') as f:
            for i in range(2, self._rows+1):
                f.write('INSERT INTO table_name VALUES(');
                data_str = ''

                for j in range(1, self._columns+1):
                   
                    if(self._ws.cell(row=i,column=j).value is None):
                        data_str += 'NULL,'

                    elif (isinstance(self._ws.cell(row=i,column=j).value, str)):
                        #Removes whitespace, linefeeds, and carriage returns
                        #TODO trimming does not seem to be working
                        data_str += "REPLACE(REPLACE(LTRIM(RTRIM('"+ self._ws.cell(row=i, column=j).value + "')),CHAR(13),''), CHAR(10), ''),"

                    else:
                        data_str += self._ws.cell(row=i, column=j).value + ","

                else:
                    f.write(data_str[:-1] + ')\n')
                
        print('Complete')

#Main Method
if __name__=='__main__':    
    ExtractExcel()
