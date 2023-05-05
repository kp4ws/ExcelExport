class ExcelExportModel:
    def __init__(self):
        self.program_type = None
        self.source_file = None
        self.destination_file = None
        self.rows = None
        self.columns = None
        self.row_start = None
        self.column_start = None
        #TODO: Delimiter (needed for create?)
    
    def generate_sql_commands(self):
        #TODO:
        #Update state
        #Perform business logic for generating sql commands
        pass
    
    def get_prefill_data(self):
        #TODO:
        pass