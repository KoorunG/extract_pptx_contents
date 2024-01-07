from datetime import datetime
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from pathlib import Path
from excel_design import design_excel
from ppt_extract import read_ppt

# 워크북
wb = Workbook()
ws: Worksheet = wb.active
# 헤더 추가
ws.append(['인덱스', 'PPT 파일명', '슬라이드 번호', '목차', '추출텍스트'])

# 0: root
root: Path = Path().cwd()

print(f'root ::: {root}')

# 1: .pptx 파일 import 경로
import_dir = root / '..' / 'import'
# 2: .xls export 경로
export_dir = root / '..' / 'export'


# import_dir 하위의 .pptx 파일을 찾아서 읽기
for i in import_dir.iterdir():
    if i.is_file() and i.suffix == ".pptx":
        read_ppt(i, ws)

# 디자인
design_excel(ws)

# 엑셀 파일명
excel_name: str = f'텍스트추출_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx'

# 저장
wb.save(export_dir / excel_name)
