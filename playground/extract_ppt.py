import re
from pathlib import Path

from openpyxl import Workbook
from openpyxl.cell import Cell
from pptx import Presentation
from pptx.slide import Slide

from playground.functions import *

# 루트 경로
root: Path = Path('/Users/koorung/Desktop/project_folder/')

# 엑셀 파일명
excel_name: str = '텍스트추출.xlsx'

# # PPT 파일명
# ppt_name: str = '제안본문_Ⅱ.전략및방법론_20240103_이재학_본문점검 크로스체크_v0.09.pptx'
# # 경로
# p: Path = root / ppt_name
# # 프레젠테이션
# prs: Presentation = Presentation(p)
# sls: list[Slide] = prs.slides

# 워크북
wb = Workbook()
ws = wb.active

# 헤더 추가
ws.append(['인덱스', 'PPT 파일명', '슬라이드 번호', '추출텍스트'])

# 전역 번호
rn: int = 1

for i in root.iterdir():
    if i.is_file() and i.suffix == ".pptx":
        print(i)
        prs: Presentation = Presentation(i)
        sls: list[Slide] = prs.slides
        for index, sl in enumerate(sls):
            # 중복제거용 set()
            lines: set[str] = set()

            # 슬라이드별 텍스트 추출하기
            extract_text_in_ppt(sl.shapes, lines)

            # 콘솔에 찍기 (검증용)
            print()
            print(f'******** {index}번째 슬라이드 ********')
            print(lines)

            # 특수문자만 있는 경우를 제외하고 엑셀에 추가
            for line in lines:
                regex = re.compile("[0-9a-zA-Zㄱ-힗]", re.MULTILINE)
                if re.match(regex, line):
                    ws.append([rn, i.name, str(index) + "번째", line])
                    rn += 1

# 디자인
tups: list[tuple[Cell]] = ws.rows

for i, tup in enumerate(tups):
    a, b, c, d = tup

    if i == 0:
        a.fill = fill()
        b.fill = fill()
        c.fill = fill()
        d.fill = fill()
        a.font = font('KoPubDotum', 'ffffff', 13)
        b.font = font('KoPubDotum', 'ffffff', 13)
        c.font = font('KoPubDotum', 'ffffff', 13)
        d.font = font('KoPubDotum', 'ffffff', 13)
    else:
        a.font = font('KoPubDotum')
        b.font = font('KoPubDotum')
        c.font = font('KoPubDotum')
        d.font = font('KoPubDotum')

    a.border = border()
    b.border = border()
    c.border = border()
    d.border = border()

auto_fit_column_size(ws, margin=10)

# 저장
wb.save(root / excel_name)
