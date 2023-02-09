from typing import Dict

from openpyxl.worksheet.worksheet import Worksheet

from configuration.column import Column as ColumnConfiguration


def write_columns_header(ws: Worksheet,
                         columns_config: Dict[str, ColumnConfiguration]):
    pass
