import tkinter as tk
from gui import ExcelExportModel, ExcelExportView, ExcelExportController
from event import EventSystem, EventChannel

def center_window(window):
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f"{width}x{height}+{x}+{y}")

def main():
    root = tk.Tk()
    
    root.geometry("500x420")
    # root.minsize(400, 300)
    # root.maxsize(1280,720)
    root.title("Excel Export")

    event_system = EventSystem()
    app_model = ExcelExportModel()
    app_view = ExcelExportView(event_system, root, master=root)
    app_controller = ExcelExportController(app_view, app_model, event_system)
    
    app_view.grid()
    center_window(root)
    app_view.mainloop()

if __name__ == '__main__':
    main()