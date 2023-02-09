from typing import Dict

from format.group import Group
from format.column import Column


def indexes(element, my_list):
    first = my_list.index(element) + 1
    last = len(my_list) - my_list[::-1].index(element)
    return first, last


class Sheet:
    def __init__(self,
                 sheet_name: str,
                 groups: Dict[str, Group],
                 columns: Dict[str, Column]):
        self.sheet_name = sheet_name
        self.groups = groups
        self.columns = columns

    def groups_limits(self):
        cols_groups = [col.group.group_name for col in self.columns.values()]
        groups_names = list(dict.fromkeys(cols_groups))
        return {
            group_name: indexes(group_name, cols_groups)
            for group_name
            in groups_names}

    def __repr__(self):
        s = []
        s.append(f'Sheet()')
        s.append(f'sheet_name: {self.sheet_name}')
        s.append(f'{len(self.groups)} groups')
        s.append(f'{len(self.columns)} columns')
        return '\n'.join(s)
