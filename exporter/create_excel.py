from io import BytesIO
from openpyxl import Workbook
from typing import Dict, List

from configuration.excel_configuration import ExcelConfiguration
from exporter.write_sheet import write_sheet


def export_excel(data: List[Dict], config: ExcelConfiguration, update_time):
    wb = Workbook()
    for n, (ws_data, ws_config) in enumerate(zip(data, config.sheets.values)):
        ws = wb.active if n == 0 else wb.create_sheet()
        write_sheet(ws, ws_data, ws_config, update_time)
    virtual_file = BytesIO()
    wb.save(virtual_file)
    virtual_file.seek(0)
