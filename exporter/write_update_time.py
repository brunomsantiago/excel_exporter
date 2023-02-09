from datetime import datetime

from openpyxl.worksheet.worksheet import Worksheet

from configuration.cell_format import CellFormat


def write_update_time(ws: Worksheet,
                      update_time: datetime,
                      update_date_format: CellFormat):
    pass
