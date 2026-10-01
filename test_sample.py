from pathlib import Path
from AsanLibraryFinder import parse_search, parse_detail

SEARCH = Path('통합자료검색 - 아산시립도서관2.html')
DETAIL = Path('통합자료검색 - 아산시립도서관..html')

s = SEARCH.read_text(encoding='utf-8', errors='ignore')
d = DETAIL.read_text(encoding='utf-8', errors='ignore')
rows = parse_search(s, 'https://lib.asan.go.kr/dls_le/', 'Princess in Black')
copies, _ = parse_detail(d)
assert len(rows) == 6
assert sum(r['상태'] == '대출가능' for r in rows) == 5
assert any(r['상호대차'] == '신청가능' for r in rows)
assert copies[0]['등록번호'] == 'UM0000019068'
assert copies[0]['반납예정일'] == '2026-10-31'
print('OK')
