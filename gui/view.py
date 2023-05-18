import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
from event import EventChannel

class ExcelExportView(tk.Frame):
    def __init__(self, event_system, window, master=None):
        super().__init__(master)
        self.event_system = event_system
        self.window = window
        self.master = master
        self.create_widgets()

    def browse_file(self, is_source_file, filetypes):
        filename = filedialog.askopenfilename(filetypes=filetypes)
        if filename:
            if is_source_file:
                self.source_file.set(filename)
            else:
                self.destination_file.set(filename)

    #TODO: Break down into smaller GUI components
    def create_widgets(self):
        self.label_program_type = ttk.Label(self, text="Program Type", anchor="w")
        self.label_program_type.grid(row=0, column=0, sticky="w", padx=3)

        self.program_type = ttk.Combobox(self, values=["Insert", "Update", "Create"], state="readonly")
        self.program_type.current(0)
        self.program_type.grid(row=0, column=1)


        self.label_source_file = ttk.Label(self, text="Source File", anchor="w")
        self.label_source_file.grid(row=1, column=0, sticky="w", padx=3)

        self.source_file = tk.StringVar() #TODO What does StringVar mean?
        self.source_file_entry = ttk.Entry(self, textvariable=self.source_file, width=23)
        self.source_file_entry.grid(row=1, column=1)

        self.button_source_file = ttk.Button(self, text="Browse", command=lambda: self.browse_file(True, [("Excel Files", "*.xlsx;*.xls")]))
        self.button_source_file.grid(row=1, column=2)
        

        self.label_destination_file = ttk.Label(self, text="Destination File", anchor="w")
        self.label_destination_file.grid(row=2, column=0, sticky="w", padx=3)

        self.destination_file = tk.StringVar()
        self.destination_file_entry = ttk.Entry(self, textvariable=self.destination_file, width=23)
        self.destination_file_entry.grid(row=2, column=1)

        self.button_destination_file = ttk.Button(self, text="Browse", command=lambda: self.browse_file(False, [("All Files", "*.*")]))
        self.button_destination_file.grid(row=2, column=2)


        #TODO: Table name field
        #TODO: Is there any easier way to add in fields so that fields below get pushed down without automatically having to update row count ?  
        # self.label_table_name = ttk.Label(self, text="Table Name", anchor="w")
        # self.label_table_name.grid(row=3, column=0, sticky="w", padx=3)

        # self.table_name = tk.StringVar()
        # self.table_name_entry = ttk.Entry(self, textvariable=self.table_name, width=23)
        # self.table_name_entry.grid(row=3, column=1)


        self.label_row_num = ttk.Label(self, text="Number of Rows", anchor="w")
        self.label_row_num.grid(row=3, column=0, sticky="w", padx=3)

        self.rows = tk.IntVar()
        self.rows_spinbox = ttk.Spinbox(self, from_=1, to=10000, textvariable=self.rows)
        self.rows_spinbox.grid(row=3, column=1)



        self.label_col_num = ttk.Label(self, text="Number of Columns", anchor="w")
        self.label_col_num.grid(row=4, column=0, sticky="w", padx=3)

        self.columns = tk.IntVar()
        self.columns_spinbox = ttk.Spinbox(self, from_=1, to=10000, textvariable=self.columns)
        self.columns_spinbox.grid(row=4, column=1)



        self.label_row_start = ttk.Label(self, text="Row Start Index", anchor="w")
        self.label_row_start.grid(row=5, column=0, sticky="w", padx=3)

        self.row_start = tk.IntVar()
        self.row_start_spinbox = ttk.Spinbox(self, from_=1, to=10000, textvariable=self.row_start)
        self.row_start_spinbox.grid(row=5, column=1)



        self.label_col_start = ttk.Label(self, text="Column Start Index", anchor="w")
        self.label_col_start.grid(row=6, column=0, sticky="w", padx=3)

        self.column_start = tk.IntVar()
        self.column_start_spinbox = ttk.Spinbox(self, from_=1, to=10000, textvariable=self.column_start)
        self.column_start_spinbox.grid(row=6, column=1)


        self.button_generate = ttk.Button(self, text="Generate SQL Commands", command=self.handle_generate)
        self.button_generate.grid(row=7, column=1, padx=3)
        self.button_generate = ttk.Button(self, text="Prefill data", command=self.handle_prefill)
        self.button_generate.grid(row=7, column=0, padx=3, pady=15)
    
    def display_success_message(self):
        messagebox.showinfo("Success", "Operation completed successfully!", parent=self.window)

    def handle_generate(self):
        data = {
            "program_type": self.program_type.get(),
            "source_file": self.source_file.get(),
            "destination_file": self.destination_file.get(),
            "rows": self.rows.get(),
            "columns": self.columns.get(),
            "row_start": self.row_start.get(),
            "column_start": self.column_start.get()
        }
        self.event_system.emit(EventChannel.GENERATE, data)

    def handle_prefill(self):
        self.event_system.emit(EventChannel.PREFILL, None)
    
    def update_widgets(self, data):
        self.program_type.set(data['program_type'].capitalize())
        self.source_file.set(data['source_file'])
        self.destination_file.set(data['destination_file'])
        self.rows.set(data['rows'])
        self.columns.set(data['columns'])
        self.row_start.set(data['row_start'])
        self.column_start.set(data['column_start'])