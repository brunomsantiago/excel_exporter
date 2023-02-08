from typing import Dict

from group import Group
from column import Column


class Sheet:
    def __init__(self,
                 sheet_name: str,
                 groups: Dict[str, Group],
                 columns: Dict[str, Column]):
        self.sheet_name = sheet_name
        self.groups = groups
        self.columns = columns

    def __repr__(self):
        s = []
        s.append(f'Sheet()')
        s.append(f'sheet_name: {self.sheet_name}')
        s.append(f'{len(self.groups)} groups')
        s.append(f'{len(self.columns)} columns')
        return '\n'.join(s)
