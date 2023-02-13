## Motivation

For many years, I have been creating excel spreadsheets with a similar look and feel. My template consists of data organized into tables with similar categories having the same background colors and a merged cell with a category title above them. The background colors are dark(ish) to provide contrast with a bold white font. I use blank columns before and after the table as borders and hide grid lines and unused columns.

(picture)

For many years I have been manually creating these spreadsheets in Excel, but when I needed to update one of them every week, I decided to create a python script to automate the process. The script I created worked well, but it was very hard to maintain. If I needed to add or remove a column, I would have to make a lot of changes, and if I needed to adapt the script to another dataset, I would have to almost start from scratch.

I promised myself that I would make a cleaner version of the excel exporter as soon as I had time. The original script was written using the pandas and xlswriter libraries, where I put the data on the sheet using pandas and formatted everything I needed using xlswriter. However, recently I needed to add hyperlinks to some cells in a particular column, and I was unable to do so with my current approach. So, I decided to completely rewrite my script to make it easier to adapt to new uses and allow for the possibility of adding hyperlinks to cells.

I don't intend to maintain this repository, nor turn it into a library. However, I believe that the new version makes it very easy to create spreadsheets with my template and may be useful to others. Additionally, I think that the code is organized well enough to be adapted to other templates or used as a reference on how to organize a script to generate excel spreadsheets with complex formatting.


## The new approach

1. Load data into a list of dictionary of lists, where each dictionary represents a sheet in the excel file. Inside the dictionary each key represents a column and its corresponding list contains the column data. Each element in the list represents a cell value.
   
   (picture)

2. If a hyperlink is required, the cell value should be represented as a tuple, with the first element being the cell value and the second element being the URL for the hyperlink.
    
    (picture)

3. Configurations for columns, groups, sheets, and the spreadsheet as a whole are stored in a YAML file.
    - Columns: Order, Variable Name (in the python dictionary), Title, Font, Width, and Group/Category it belongs to.
    - Groups: Name and Background Color
    - Sheets: Name, Group(s), and Column(s)
    - Spreadsheet: Cell formats and Sheets.

    (picture)

## How to use

### 1. Prepare the Data
It's necessary to properly prepare your data before using the exporter.
```python
data = [
    {
        "employee_id": [(1001, 'https://employee.my_company.com/id/1001'),
                        (1002, 'https://employee.my_company.com/id/1002')],
        "employee_name": ["John Doe", "Jane Doe"],
        "department": ["Sales South Area", "Sales West Area"],
        "job_title": ["Manager", "Associate"]
    },
    {
        "employee_id": [1001, 1002],
        "sales_current_year": [100000, 120000],
        "sales_last_year": [90000, 110000]
    }
]
```

- If your data originally comes from a list of dictionaries, you can use the `utils` module to convert it.
```python
# to fiil later: data conversion block of code
```
- If your data is in a pandas dataframe, a similar conversion process applies.
```python
# to fiil later: data conversion block of code
```

### 2. Load the YAML
Before using the exporter, you'll need to load the YAML configuration file that you'll be using to export your data.
```python
from configuration.load_config import load_config

config = load_config('example_config.yaml')
```

### 3. Retrieve the Current Date and Time
This exporter was created for periodically updated spreadsheets. So it also important to include the time and date of the dataset. These can be retrieved usion datetime module.
```python
from datetime import datetime

update_time = datetime.now()
```

### 3. Run the Exporter
Once your data is prepared and the YAML is loaded, you're ready to run the exporter. Simply pass your data, configuration and update time as arguments to the exporter, and the export process will begin. It will return a virtual file (BytesIO)
```python
from exporter.create_excel import export_excel

virtual_file = export_excel(data, config, update_time)
```
### 4. Save or upload the file
Once the export process is complete, you can choose to either save the virtual file as a real file on your local machine, or upload it to a remote server using an API. Here is an example of how to save the file locally:
```python
with open('example_output.xlsx', 'wb') as file:
    file.write(virtual_file.getbuffer())
```