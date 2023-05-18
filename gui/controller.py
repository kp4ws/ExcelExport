from event import EventChannel

class ExcelExportController:
    def __init__(self, view, model, event_system):
        self.view = view
        self.model = model
        self.event_system = event_system
        self.event_system.subscribe(EventChannel.GENERATE, self.generate_sql_commands)
        self.event_system.subscribe(EventChannel.PREFILL, self.prefill_data)

    def generate_sql_commands(self, data):
        self.model.generate_sql_commands(data)
        self.view.display_success_message()
    
    def prefill_data(self, data):
        data = self.model.get_prefill_data()
        self.view.update_widgets(data)