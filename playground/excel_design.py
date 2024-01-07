from openpyxl.cell import Cell
from openpyxl.styles import Font, Border, PatternFill
from openpyxl.styles import Side
from openpyxl.worksheet.worksheet import Worksheet


def design_excel(ws: Worksheet):
    rows: list[tuple[Cell]] = ws.rows

    # 폰트설정
    font_style: str = 'KoPub돋움체 Medium'

    for i, tup in enumerate(rows):
        for el in tup:
            el.fill = __base_fill('000000') if i == 0 else __base_fill('ffffff')
            el.font = __base_font(font_style, 'ffffff', 13) if i == 0 else __base_font(font_style, '000000', 11)
            el.border = __base_border()
    __auto_fit_column_size(ws, margin=10)


def __auto_fit_column_size(worksheet, columns=None, margin=2):
    for i, column_cells in enumerate(worksheet.columns):
        is_ok = False
        if columns is None or (isinstance(columns, list) and i in columns):
            is_ok = True

        if is_ok:
            length = max(len(str(cell.value)) for cell in column_cells)
            worksheet.column_dimensions[column_cells[0].column_letter].width = length + margin

    return worksheet


def __base_border(border_style='thin', rgb='000000'):
    return Border(left=Side(border_style, rgb),
                  right=Side(border_style, rgb),
                  top=Side(border_style, rgb),
                  bottom=Side(border_style, rgb))


def __base_font(name, rgb, size):
    return Font(name=name, color=rgb, size=size)


def __base_fill(rgb: str):
    return PatternFill(start_color=rgb, end_color=rgb, fill_type='solid')
