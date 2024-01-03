from openpyxl.styles import Font, Border, PatternFill
from openpyxl.styles import Side
from pptx.enum.shapes import MSO_SHAPE_TYPE


# PPT에서 텍스트 추출하기
def extract_text_in_ppt(shapes, texts: set[str]):
    for shape in shapes:
        # 1. 텍스트 프레임인 경우
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    texts.add(run.text)
                    # print(run.text, end="")
        # 2. 표인 경우
        elif shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    for cell_paragraph in cell.text_frame.paragraphs:
                        for cell_run in cell_paragraph.runs:
                            texts.add(cell_run.text)
                            # print(cell_run.text, end="")

        # 3. 그룹인경우 재귀
        elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            extract_text_in_ppt(shape.shapes, texts)


def auto_fit_column_size(worksheet, columns=None, margin=2):
    for i, column_cells in enumerate(worksheet.columns):
        is_ok = False
        if columns == None:
            is_ok = True
        elif isinstance(columns, list) and i in columns:
            is_ok = True

        if is_ok:
            length = max(len(str(cell.value)) for cell in column_cells)
            worksheet.column_dimensions[column_cells[0].column_letter].width = length + margin

    return worksheet


def border(border_style='thin', color='000000'):
    return Border(left=Side(border_style, color),
                  right=Side(border_style, color),
                  top=Side(border_style, color),
                  bottom=Side(border_style, color))


def font(name, color='000000', size=11):
    return Font(name=name, color=color, size=size)


def fill():
    return PatternFill(start_color='000000', end_color='000000', fill_type='solid')
