from datetime import datetime
from openpyxl import Workbook
from playground.functions import *

# 워크북
wb = Workbook()
ws = wb.active
# 헤더 추가
ws.append(['인덱스', 'PPT 파일명', '슬라이드 번호', '추출텍스트'])

# 전역 번호
rn: int = 1

# 루트 경로
root: Path = Path('C:/Users/jaehak/Desktop/회사/00. 제안작업/1. 본문/★ 본문점검/☆ 본문점검 제안요청서 매칭/추출대상 PPT')

# 루트 하위의 .pptx 파일을 찾아서 ppt 읽기
for i in root.iterdir():
    if i.is_file() and i.suffix == ".pptx":
        read_ppt(rn, i, ws)

# 디자인
design_excel(ws)

# 엑셀 파일명
excel_name: str = f'텍스트추출_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx'

# 저장
wb.save(root / excel_name)
