import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

class ExcelExportView(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.create_widgets()

    #TODO: Break down into smaller GUI components
    def create_widgets(self):
        self.label = ttk.Label(self, text="Program Type", anchor="w")
        self.label.grid(row=0, column=0, sticky="w", padx=3)

        self.program_type = ttk.Combobox(self, values=["Insert", "Update", "Create"], state="readonly")
        self.program_type.current(0)
        self.program_type.grid(row=0, column=1)

        self.source_file = tk.StringVar()
        self.source_file_entry = ttk.Entry(self, textvariable=self.source_file, width=23)
        self.source_file_entry.grid(row=1, column=1)

        self.source_file_button = ttk.Button(self, text="Browse", command=lambda: self.browse_file(True, [("Excel Files", "*.xlsx;*.xls")]))
        self.source_file_button.grid(row=1, column=2)

        self.label = ttk.Label(self, text="Source File", anchor="w")
        self.label.grid(row=1, column=0, sticky="w", padx=3)

        self.label = ttk.Label(self, text="Destination File", anchor="w")
        self.label.grid(row=2, column=0, sticky="w", padx=3)

        self.destination_file = tk.StringVar()
        self.destination_file_entry = ttk.Entry(self, textvariable=self.destination_file, width=23)
        self.destination_file_entry.grid(row=2, column=1)

        self.destination_file_button = ttk.Button(self, text="Browse", command=lambda: self.browse_file(False, [("SQL Files", "*.sql")]))
        self.destination_file_button.grid(row=2, column=2)

        self.label = ttk.Label(self, text="Number of Rows", anchor="w")
        self.label.grid(row=3, column=0, sticky="w", padx=3)

        self.rows = tk.IntVar()
        self.rows_spinbox = ttk.Spinbox(self, from_=1, to=10000, textvariable=self.rows)
        self.rows_spinbox.grid(row=3, column=1)

        self.label = ttk.Label(self, text="Number of Columns", anchor="w")
        self.label.grid(row=4, column=0, sticky="w", padx=3)

        self.columns = tk.IntVar()
        self.columns_spinbox = ttk.Spinbox(self, from_=1, to=10000, textvariable=self.columns)
        self.columns_spinbox.grid(row=4, column=1)

        self.label = ttk.Label(self, text="Row Start Index", anchor="w")
        self.label.grid(row=5, column=0, sticky="w", padx=3)

        self.row_start = tk.IntVar()
        self.row_start_spinbox = ttk.Spinbox(self, from_=1, to=10000, textvariable=self.row_start)
        self.row_start_spinbox.grid(row=5, column=1)

        self.label = ttk.Label(self, text="Column Start Index", anchor="w")
        self.label.grid(row=6, column=0, sticky="w", padx=3)

        self.column_start = tk.IntVar()
        self.column_start_spinbox = ttk.Spinbox(self, from_=1, to=10000, textvariable=self.column_start)
        self.column_start_spinbox.grid(row=6, column=1)

        self.generate_button = ttk.Button(self, text="Generate SQL Commands")
        self.generate_button.grid(row=7, column=1, padx=3)
        self.generate_button = ttk.Button(self, text="Prefill data")
        self.generate_button.grid(row=7, column=0, padx=3, pady=15)

    def browse_file(self, is_source_file, filetypes):
        filename = filedialog.askopenfilename(filetypes=filetypes)
        if filename:
            if is_source_file:
                self.source_file.set(filename)
            else:
                self.destination_file.set(filename)