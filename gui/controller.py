class ExcelExportController:
    def __init__(self, view, model):
        self.view = view
        self.model = model
        self.view.generate_button.config(command=self.generate_sql_commands)
        self.view.prefill_button.config(command=self.prefill_data)

    def generate_sql_commands(self, data):
        # self.model.generate_sql_commands(data)
        pass
    
    def prefill_data(self):
        #TODO:
        #data = self.model.get_prefill_data()
        #self.view.update_fields(data)
        pass