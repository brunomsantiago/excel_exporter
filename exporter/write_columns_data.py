from typing import Dict, List

from openpyxl.worksheet.worksheet import Worksheet

from configuration.column import Column as ColumnConfiguration


def write_columns_data(ws: Worksheet,
                       ws_data: Dict[str, List],
                       columns_config: Dict[str, ColumnConfiguration]):
    pass
