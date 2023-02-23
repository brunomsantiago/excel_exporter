from datetime import datetime
from excel_exporter.configuration.load_config import load_config
from excel_exporter.exporter.create_excel import export_excel

# Prepare the Data
data = [
    {
        'employee_id': [
            (1001, 'https://employee.my_company.com/id/1001'),
            (1002, 'https://employee.my_company.com/id/1002'),
        ],
        'employee_name': ['John Doe', 'Jane Doe'],
        'department': ['Sales South Area', 'Sales West Area'],
        'job_title': ['Manager', 'Associate'],
    },
    {
        'employee_id': [1001, 1002],
        'sales_current_year': [100000, 120000],
        'sales_last_year': [90000, 110000],
    },
]
# Load the YAML
config = load_config('example_config.yaml')
# Retrieve the Current Date and Time
update_time = datetime.now()
# Run the Exporter
virtual_file = export_excel(data, config, update_time)
# Save the file
with open(config.file_name, 'wb') as file:
    file.write(virtual_file.getbuffer())
