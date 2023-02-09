from typing import Dict

from configuration.cell_format import CellFormat
from configuration.sheet import Sheet


class ExcelConfiguration:
    def __init__(self,
                 file_name: str,
                 cell_formats: Dict[str, CellFormat],
                 sheets: Dict[str, Sheet]):
        self.file_name = file_name
        self.cell_formats = cell_formats
        self.sheets = sheets

    def __repr__(self, n_indent=2):
        indent = ' ' * n_indent
        s = []
        s.append(f'ExcelConfiguration()')
        s.append(f'{indent}file_name: {self.file_name}')
        s.append(f'{indent}{len(self.cell_formats)} cell_formats')
        s.append(f'{indent}{len(self.sheets)} sheets ')
        return '\n'.join(s)