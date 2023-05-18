import json
from extract import insert, create, update

class ExcelExportModel:
    def update_prefill_data(self, data):
        with open('resources/config.json', 'w') as file:
            json.dump(data, file)

    def get_prefill_data(self):
        with open('resources/config.json', 'r') as file:
            data = json.load(file)
        return data

    def generate_sql_commands(self, data):
        program_type = data['program_type']

        if program_type == 'Insert':
            sql_commands = insert.generate_insert_sql_commands(data)
        elif program_type == 'Create':
            sql_commands = create.generate_create_sql_commands(data)
        elif program_type == 'Update':
            sql_commands = update.generate_update_sql_commands(data)
        else:
            raise ValueError("Invalid program type" + program_type)
        
        destination_file = data["destination_file"]
        self.save_sql_commands(sql_commands, destination_file)
        
        self.update_prefill_data(data)

    def save_sql_commands(self, sql_commands, destination_file):
        with open(destination_file, 'w') as file:
            for command in sql_commands:
                file.write(command)
                file.write('\n')