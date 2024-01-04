import re

from openpyxl.cell import Cell
from openpyxl.styles import Font, Border, PatternFill
from openpyxl.styles import Side
from openpyxl.worksheet.worksheet import Worksheet
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.slide import Slide
from pathlib import Path


# PPT에서 텍스트 추출하기
def extract_text_in_ppt(shapes) -> set[str]:
    # 중복제거용 set()
    lines: set[str] = set()

    for shape in shapes:
        # 1. 텍스트 프레임인 경우
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    lines.add(run.text)
                    # print(run.text, end="")
        # 2. 표인 경우
        elif shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    for cell_paragraph in cell.text_frame.paragraphs:
                        for cell_run in cell_paragraph.runs:
                            lines.add(cell_run.text)
                            # print(cell_run.text, end="")

        # 3. 그룹인경우 재귀
        elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            extract_text_in_ppt(shape.shapes)

    return lines


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


def base_border(border_style='thin', rgb='000000'):
    return Border(left=Side(border_style, rgb),
                  right=Side(border_style, rgb),
                  top=Side(border_style, rgb),
                  bottom=Side(border_style, rgb))


def base_font(name, rgb, size):
    return Font(name=name, color=rgb, size=size)


def base_fill(rgb: str):
    return PatternFill(start_color=rgb, end_color=rgb, fill_type='solid')


def design_excel(ws: Worksheet):
    rows: list[tuple[Cell]] = ws.rows

    # 폰트설정
    font_style: str = 'KoPub돋움체 Medium'

    for i, tup in enumerate(rows):
        for el in tup:
            el.fill = base_fill('000000') if i == 0 else base_fill('ffffff')
            el.font = base_font(font_style, 'ffffff', 13) if i == 0 else base_font(font_style, '000000', 11)
            el.border = base_border()
    auto_fit_column_size(ws, margin=10)


def read_ppt(n: int, path: Path, ws: Worksheet):
    prs: Presentation = Presentation(path)
    sls: list[Slide] = prs.slides
    for i, sl in enumerate(sls):
        print(f'[{path.name}] :: {i}번째 슬라이드 읽음')
        
        # 슬라이드별 텍스트 추출하기 & 특수문자만 있는 경우를 제외하고 엑셀에 추가
        [append_row(i, n, text, path.name, ws) for text in extract_text_in_ppt(sl.shapes)]


def append_row(slide_index: int, index: int, text: str, file_name: str, ws: Worksheet):
    regex = re.compile("[0-9a-zA-Zㄱ-힗]", re.MULTILINE)
    if re.match(regex, text):
        ws.append([index, file_name, str(slide_index) + "번째", text])
        index += 1
