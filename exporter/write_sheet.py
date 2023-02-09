from typing import Dict

from openpyxl.worksheet.worksheet import Worksheet

from configuration.sheet import Sheet as SheetConfiguration
from exporter.apply_sheet_configs import apply_sheet_configs
from exporter.write_groups_header import write_groups_header
from exporter.write_update_time import write_update_time
from exporter.write_columns_data import write_columns_data
from exporter.write_columns_header import write_columns_header


def write_sheet(ws: Worksheet,
                ws_data: Dict,
                ws_config: SheetConfiguration,
                update_time,
                ):
    write_update_time(ws, update_time, ws_config.update_date_format)
    write_groups_header(ws, ws_config.groups, ws_config.groups_limits())
    write_columns_header(ws, ws_config.columns)
    write_columns_data(ws, ws_data, ws_config.columns)
    apply_sheet_configs(ws, ws_config)