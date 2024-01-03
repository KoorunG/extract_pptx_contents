from pptx import Presentation
from pptx.slide import Slide
from pptx.enum.shapes import MSO_SHAPE_TYPE

import openpyxl
from pathlib import Path

# 루트 경로
root: Path = Path('/Users/koorung/Desktop/project_folder/')

# 파일명
file_name: str = '제안본문_Ⅱ.전략및방법론_20240103_이재학_본문점검 크로스체크_v0.09.pptx'

# 경로
p: Path = root / file_name


# PPT에서 텍스트 추출하기
def extract_text_in_ppt(shapes):
    for shape in shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    print(run.text)
        elif shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    for cell_paragraph in cell.text_frame.paragraphs:
                        for cell_run in cell_paragraph.runs:
                            print(cell_run.text)
        elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            extract_text_in_ppt(shape.shapes)


# 프레젠테이션
prs: Presentation = Presentation(p)
sls: list[Slide] = prs.slides
for sl in sls:
    extract_text_in_ppt(sl.shapes)
