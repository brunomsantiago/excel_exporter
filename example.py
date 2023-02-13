from datetime import datetime
from configuration.load_config import load_config
from exporter.create_excel import export_excel

data = [
    {
        "employee_id": [(1001, 'https://employee.my_company.com/id/1001'),
                        (1002, 'https://employee.my_company.com/id/1002')],
        "employee_name": ["John Doe", "Jane Doe"],
        "department": ["Marketing", "Sales"],
        "job_title": ["Manager", "Associate"]
    },
    {
        "employee_id": [1001, 1002],
        "sales_current_year": [100000, 120000],
        "sales_last_year": [90000, 110000]
    }
]


update_time = datetime.utcnow()
config = load_config('example_config.yaml')
virtual_file = export_excel(data, config, update_time)
with open('example_output.xlsx', 'wb') as f:
    f.write(virtual_file.getbuffer())
