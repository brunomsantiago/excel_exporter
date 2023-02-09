from datetime import datetime
from configuration.load_config import load_config
from exporter.create_excel import export_excel

data = [
    {
        "employee_id": [(1001, 'https://www.google.com'),
                        (1002, 'https://www.bing.com'),
                        (1003, 'https://www.yahoo.com'),
                        (1004, 'https://duckduckgo.com'),
                        (1005, 'https://tineye.com/')],
        "employee_name": ["John Doe", "Jane Doe", "Jim Smith", "Amy Johnson", "Michael Brown"],
        "department": ["Marketing", "Sales", "IT", "HR", "Finance"],
        "job_title": ["Manager", "Associate", "Developer", "Analyst", "Director"],
        "date_of_birth": ["1980-01-01", "1985-02-02", "1990-03-03", "1995-04-04", "2000-05-05"],
        "phone_number": ["+1-111-111-1111", "+1-222-222-2222", "+1-333-333-3333", "+1-444-444-4444", "+1-555-555-5555"],
        "address": ["123 Main St", "456 Elm St", "789 Oak St", "246 Cedar Ave", "369 Maple Rd"]
    },
    {
        "employee_id": [1001, 1002, 1003, 1004, 1005],
        "employee_name": ["John Doe", "Jane Doe", "Jim Smith", "Amy Johnson", "Michael Brown"],
        "sales_current_year": [100000, 120000, 150000, 140000, 110000],
        "scorecard_current_year": [90, 92, 95, 87, 88],
        "sales_last_year": [90000, 110000, 130000, 120000, 100000],
        "scorecard_last_year": [85, 88, 90, 83, 82]
    }
]


update_time = datetime.utcnow()
config = load_config('z_config1.yaml')
virtual_file = export_excel(data, config, update_time)
with open('former_virtual_file.xlsx', 'wb') as f:
    f.write(virtual_file.getbuffer())
