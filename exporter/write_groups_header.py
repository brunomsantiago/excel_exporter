from typing import Dict, Tuple

from openpyxl.worksheet.worksheet import Worksheet

from configuration.group import Group as GroupConfiguration


def write_groups_header(ws: Worksheet,
                        groups_config: Dict[str, GroupConfiguration],
                        groups_limits: Dict[str, Tuple]):
    pass
