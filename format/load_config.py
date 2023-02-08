import yaml

from format.cell_format import CellFormat
from format.column import Column
from format.excel_configuration import ExcelConfiguration
from format.group import Group
from format.sheet import Sheet


def parse_yaml(yaml_data) -> ExcelConfiguration:

    file_name = yaml_data["file_name"]

    cell_formats = {}
    for cell_format_data in yaml_data["cell_formats"]:
        cell_format = CellFormat(cell_format_data["format_name"],
                                 cell_format_data["font_size"],
                                 cell_format_data["horizontal_alignment"],
                                 cell_format_data["vertical_alignment"],
                                 cell_format_data["line_break"])
        cell_formats[cell_format.format_name] = cell_format

    sheets = {}
    for sheet_data in yaml_data["sheets"]:
        groups = {}
        for group_data in sheet_data["groups"]:
            group = Group(group_data["group_name"],
                          group_data["background_color"])
            groups[group.group_name] = group

        columns = {}
        for column_data in sheet_data["columns"]:
            cell_format = cell_formats.get(column_data["cell_format"], None)
            if not cell_format:
                raise ValueError(
                    f"Cell format {column_data['cell_format']} not found")

            group = groups.get(column_data["group"], None)
            if not group:
                raise ValueError(f"Group {column_data['group']} not found")

            column = Column(column_data["column_name"],
                            column_data["variable_name"],
                            cell_format,
                            group,
                            column_data["column_width"])
            columns[column.column_name] = column

        sheet = Sheet(sheet_data["sheet_name"], groups, columns)
        sheets[sheet_data["sheet_name"]] = sheet
    return ExcelConfiguration(file_name, cell_formats, sheets)


def load_config(file_path: str) -> ExcelConfiguration:
    with open(file_path, "r") as file:
        yaml_data = yaml.safe_load(file)
    return parse_yaml(yaml_data)
