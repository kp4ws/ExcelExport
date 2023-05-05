import tkinter as tk

from view import ExcelExportView
from model import ExcelExportModel
from controller import ExcelExportController

def main():
    root = tk.Tk()
    root.geometry("480x360")
    root.minsize(400, 300)
    root.maxsize(480, 360)
    root.title("Excel Export")
    app_model = ExcelExportModel()
    app_view = ExcelExportView(master=root)
    #app_controller = ExcelExportController(app_view, app_model)
    app_view.grid()
    app_view.mainloop()

if __name__ == '__main__':
    main()