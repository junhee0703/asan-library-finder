import re, math, threading, webbrowser
from urllib.parse import urljoin
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import requests
from bs4 import BeautifulSoup

BASE = 'https://lib.asan.go.kr/dls_le/'
SEARCH_URL = urljoin(BASE, 'index.php')

# Tangjeong-myeon reference point. Distance shown by the app is straight-line distance.
TANGJEONG_REF = (36.79448, 127.05510)

LIBRARY_ADDRESSES = {
    '탕정온샘도서관': '충청남도 아산시 탕정면 한내로 209',
    '배방월천도서관': '충청남도 아산시 배방읍 월천5길 13',
    '배방도서관': '충청남도 아산시 배방읍 모산로 182-14',
    '아산중앙도서관': '충청남도 아산시 남부로 229',
    '꿈샘어린이청소년도서관': '충청남도 아산시 시민로 500',
    '음봉어울샘도서관': '충청남도 아산시 음봉면 음봉로 515-36',
    '둔포도서관': '충청남도 아산시 둔포면 둔포중앙로161번길 27-6',
    '신창도서관': '충청남도 아산시 신창면 남성길 99',
    '남산도서관': '충청남도 아산시 남산로 10',
}

# Fallback coordinates. If a library is missing, the program tries Nominatim once at runtime.
LIBRARY_COORDS = {
    '탕정온샘도서관': (36.79448, 127.05510),
    '배방월천도서관': (36.77060, 127.07830),
    '배방도서관': (36.77391, 127.05784),
    '아산중앙도서관': (36.77034, 127.00573),
    '음봉어울샘도서관': (36.8440, 127.0670),
}

SESSION = requests.Session()
SESSION.headers.update({
    'User-Agent': 'AsanLibraryFinder/3.0 (Windows; library-search utility)',
    'Accept-Language': 'ko-KR,ko;q=0.9,en;q=0.8',
})


def clean(s):
    return re.sub(r'\s+', ' ', s or '').strip()


def haversine_km(a, b):
    lat1, lon1 = map(math.radians, a)
    lat2, lon2 = map(math.radians, b)
    dlat, dlon = lat2 - lat1, lon2 - lon1
    h = math.sin(dlat / 2) ** 2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2) ** 2
    return 6371.0088 * 2 * math.asin(math.sqrt(h))


def geocode_library(name):
    if name in LIBRARY_COORDS:
        return LIBRARY_COORDS[name]
    address = LIBRARY_ADDRESSES.get(name)
    if not address:
        return None
    try:
        r = SESSION.get(
            'https://nominatim.openstreetmap.org/search',
            params={'q': address, 'format': 'jsonv2', 'limit': 1},
            timeout=10,
        )
        r.raise_for_status()
        data = r.json()
        if data:
            coord = (float(data[0]['lat']), float(data[0]['lon']))
            LIBRARY_COORDS[name] = coord
            return coord
    except Exception:
        return None
    return None


def search_dls(query):
    # The supplied search HTML shows the real form action and a per-page CSRF token.
    # First load the search form, then submit the same GET fields the site uses.
    form_url = SEARCH_URL
    first = SESSION.get(form_url, params={'mod': 'wdDataSearch', 'act': 'dataSearchForm'}, timeout=25)
    first.raise_for_status()
    first.encoding = first.apparent_encoding or 'utf-8'
    soup = BeautifulSoup(first.text, 'html.parser')
    form = soup.find('form', attrs={'name': 'headSearchForm'}) or soup.find('form', action=re.compile('searchIList'))
    token = ''
    if form:
        inp = form.find('input', attrs={'name': 'WD_CSRF_TOKEN'})
        token = inp.get('value', '') if inp else ''

    params = {
        'mod': 'wdDataSearch',
        'act': 'searchIList',
        'WD_CSRF_TOKEN': token,
        'preWord': query,
        'preItem': 'total',
        'preCond': '',
        'deSearch': '2',
        'item': 'total',
        'word': query,
        'listNum': '20',
    }
    r = SESSION.get(SEARCH_URL, params=params, timeout=25)
    r.raise_for_status()
    r.encoding = r.apparent_encoding or 'utf-8'
    return r.text, r.url


def _field_from_li(li, label):
    for strong in li.find_all('strong'):
        if label not in clean(strong.get_text(' ', strip=True)):
            continue
        chunks = []
        for sib in strong.next_siblings:
            if getattr(sib, 'name', None) in ('strong', 'em'):
                break
            chunks.append(sib.get_text(' ', strip=True) if hasattr(sib, 'get_text') else str(sib))
        return clean(' '.join(chunks))
    return ''


def parse_search(html, base_url, query=''):
    soup = BeautifulSoup(html, 'html.parser')
    rows = []
    for box in soup.select('div.divList > div.list'):
        links = box.select("a[href*='act=searchResultDetail']")
        link = next((a for a in links if clean(a.get_text(' ', strip=True))), None)
        if not link:
            continue

        href = urljoin(base_url, link.get('href', ''))
        title = clean(link.get_text(' ', strip=True))
        ul = box.find('ul', recursive=False)
        lis = ul.find_all('li', recursive=False) if ul else []

        author = publisher = year = callno = library = room = ''
        for li in lis:
            author = author or _field_from_li(li, '저자')
            publisher = publisher or _field_from_li(li, '발행처')
            y = _field_from_li(li, '발행년')
            year = year or ((re.search(r'\d{4}', y) or [''])[0])
            callno = callno or _field_from_li(li, '청구기호')
            if '소장기관' in clean(li.get_text(' ', strip=True)):
                spans = li.select('span.fb')
                if len(spans) >= 2:
                    library = clean(spans[0].get_text(' ', strip=True))
                    room = clean(spans[1].get_text(' ', strip=True))

        all_text = clean(box.get_text(' ', strip=True))
        status = ''
        for strong in box.select('ol strong'):
            t = clean(strong.get_text(' ', strip=True))
            if t:
                if '대출가능' in t:
                    status = '대출가능'
                elif '대출중' in t:
                    status = '대출중'
                elif '예약중' in t:
                    status = '예약중'
                elif not status and ('자료실내열람' in t or '열람' in t):
                    status = t
        if not status:
            if '대출가능' in all_text:
                status = '대출가능'
            elif '대출중' in all_text:
                status = '대출중'

        interlibrary = '신청가능' if '상호대차 신청가능' in all_text else ''
        reservation = '가능' if '예약가능' in all_text else ('불가' if '예약불가' in all_text else '')

        rows.append({
            '검색어': query,
            '제목': title,
            '저자': author,
            '발행처': publisher,
            '발행년': year,
            '소장기관': library,
            '자료실': room,
            '청구기호': callno,
            '상태': status,
            '반납예정일': '',
            '등록번호': '',
            '낱권': '',
            '예약': reservation,
            '상호대차': interlibrary,
            '딸림자료': '',
            '거리_km': None,
            '상세URL': href,
        })
    return rows


def parse_detail(html):
    soup = BeautifulSoup(html, 'html.parser')
    copies = []
    table = None
    for t in soup.find_all('table'):
        heads = [clean(x.get_text(' ', strip=True)) for x in t.find_all('th')]
        if '자료상태' in heads and '반납예정일' in heads:
            table = t
            break

    if table:
        headers = [clean(x.get_text(' ', strip=True)) for x in table.find_all('th')]
        for tr in table.find_all('tr'):
            tds = tr.find_all('td', recursive=False)
            if not tds:
                continue
            vals = [clean(td.get_text(' ', strip=True)) for td in tds]
            vals += [''] * max(0, len(headers) - len(vals))
            row = dict(zip(headers, vals))
            first = tds[0]
            first_text = clean(first.get_text(' ', strip=True))
            m = re.search(r'\b[A-Z]{2}\d{8,}\b', first_text)
            reg = m.group(0) if m else ''
            copies.append({
                '등록번호': reg,
                '낱권': row.get('낱권', ''),
                '자료실': row.get('자료실 / 청구기호', ''),
                '상태': row.get('자료상태', ''),
                '반납예정일': row.get('반납예정일', ''),
                '예약': row.get('예약', ''),
                '상호대차': row.get('상호대차', ''),
            })

    # The supplied detail sample does not contain a visible 딸림자료 section.
    # Scan headings/labels only when they actually exist on a future detail page.
    supp = ''
    for node in soup.find_all(string=re.compile(r'딸림자료|부록자료|부록')):
        txt = clean(node.parent.parent.get_text(' ', strip=True) if node.parent and node.parent.parent else str(node))
        if txt and len(txt) < 500:
            supp = txt
            break
    return copies, supp


def choose_copy(row, copies):
    if not copies:
        return None
    # Search-result library/room/call-number are used to identify the exact copy.
    best = None
    for c in copies:
        ct = c.get('자료실', '')
        score = 0
        if row.get('자료실') and row['자료실'] in ct:
            score += 4
        if row.get('청구기호') and row['청구기호'] in ct:
            score += 3
        if row.get('상태') and row['상태'] in c.get('상태', ''):
            score += 1
        if best is None or score > best[0]:
            best = (score, c)
    return best[1]


def enrich_row(row, detail_cache):
    url = row['상세URL']
    if url not in detail_cache:
        r = SESSION.get(url, timeout=25)
        r.raise_for_status()
        r.encoding = r.apparent_encoding or 'utf-8'
        detail_cache[url] = parse_detail(r.text)
    copies, supp = detail_cache[url]
    c = choose_copy(row, copies)
    if c:
        row['등록번호'] = c.get('등록번호', '')
        row['낱권'] = c.get('낱권', '')
        row['반납예정일'] = c.get('반납예정일', '')
        if c.get('상태'):
            row['상태'] = c['상태']
        if c.get('상호대차') and not row['상호대차']:
            row['상호대차'] = c['상호대차']
        if c.get('예약') and not row['예약']:
            row['예약'] = c['예약']
    row['딸림자료'] = supp
    return row


def run_queries(queries, progress=None):
    all_rows = []
    detail_cache = {}
    geo_cache = {}
    for qi, q in enumerate(queries, 1):
        if progress:
            progress(f'검색 중: {q} ({qi}/{len(queries)})')
        html, base_url = search_dls(q)
        rows = parse_search(html, base_url, q)
        for i, row in enumerate(rows, 1):
            lib = row['소장기관']
            if lib not in geo_cache:
                geo_cache[lib] = geocode_library(lib)
            coord = geo_cache[lib]
            if coord:
                row['거리_km'] = round(haversine_km(TANGJEONG_REF, coord), 2)
            try:
                enrich_row(row, detail_cache)
            except Exception as e:
                row['상세오류'] = str(e)
            if progress:
                progress(f'{q}: {i}/{len(rows)}개 소장본 확인')
            all_rows.append(row)

    # Within each search term: borrowable first, then distance. Keep unavailable copies separate in UI.
    def key(r):
        borrowable = 0 if r.get('상태') == '대출가능' else 1
        dist = r.get('거리_km') if r.get('거리_km') is not None else 9999
        return (r.get('검색어', ''), borrowable, dist, r.get('소장기관', ''), r.get('자료실', ''))
    return sorted(all_rows, key=key)


def save_xlsx(rows, path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment
    wb = Workbook()
    ws = wb.active
    ws.title = '전체결과'
    cols = ['검색어','제목','저자','발행처','발행년','소장기관','거리_km','자료실','청구기호','등록번호','낱권','상태','반납예정일','예약','상호대차','딸림자료','상세URL']
    ws.append(cols)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal='center')
    for r in rows:
        ws.append([r.get(c, '') for c in cols])
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    for col, width in {'A':20,'B':44,'C':34,'D':24,'E':10,'F':18,'G':10,'H':28,'I':24,'J':18,'K':12,'L':12,'M':14,'N':10,'O':14,'P':35,'Q':85}.items():
        ws.column_dimensions[col].width = width

    for name, subset in [('대출가능', [r for r in rows if r.get('상태') == '대출가능']),
                         ('대출중_기타', [r for r in rows if r.get('상태') != '대출가능'])]:
        sh = wb.create_sheet(name)
        sh.append(cols)
        for c in sh[1]:
            c.font = Font(bold=True)
        for r in subset:
            sh.append([r.get(c, '') for c in cols])
        sh.freeze_panes = 'A2'
        sh.auto_filter.ref = sh.dimensions
        for col, width in {'A':20,'B':44,'C':34,'D':24,'E':10,'F':18,'G':10,'H':28,'I':24,'J':18,'K':12,'L':12,'M':14,'N':10,'O':14,'P':35,'Q':85}.items():
            sh.column_dimensions[col].width = width
    wb.save(path)


DISPLAY_COLS = ['검색어','제목','소장기관','거리(km)','자료실','청구기호','상태','반납예정일','상호대차','딸림자료']


def fill_tree(tree, rows):
    for item in tree.get_children():
        tree.delete(item)
    for r in rows:
        vals = [
            r.get('검색어',''), r.get('제목',''), r.get('소장기관',''), r.get('거리_km',''),
            r.get('자료실',''), r.get('청구기호',''), r.get('상태',''), r.get('반납예정일',''),
            r.get('상호대차',''), r.get('딸림자료','') or '표시없음'
        ]
        tree.insert('', 'end', values=vals)


class App:
    def __init__(self, root):
        self.root = root
        root.title('아산시립도서관 책 찾기')
        root.geometry('1320x760')

        top = ttk.Frame(root, padding=10)
        top.pack(fill='x')
        ttk.Label(top, text='도서 제목 — 한 줄에 하나씩').pack(anchor='w')
        self.input = tk.Text(top, height=6)
        self.input.pack(fill='x', pady=5)
        self.input.insert('1.0', 'Princess in Black\nJigsaw Jones Mystery\nMarvin Redpost\nThe Zack Files\nJudy Moody')

        bar = ttk.Frame(top)
        bar.pack(fill='x')
        self.btn = ttk.Button(bar, text='검색 시작', command=self.start)
        self.btn.pack(side='left')
        self.export_btn = ttk.Button(bar, text='엑셀 저장', command=self.export)
        self.export_btn.pack(side='left', padx=6)
        ttk.Button(bar, text='검색 페이지 열기', command=self.open_site).pack(side='left')
        self.status = ttk.Label(bar, text='대기 중')
        self.status.pack(side='left', padx=12)

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=(0,10))
        self.available_tab = ttk.Frame(self.notebook)
        self.other_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.available_tab, text='대출가능')
        self.notebook.add(self.other_tab, text='대출중 / 기타')
        self.trees = []
        for tab in (self.available_tab, self.other_tab):
            frame = ttk.Frame(tab)
            frame.pack(fill='both', expand=True)
            tree = ttk.Treeview(frame, columns=DISPLAY_COLS, show='headings')
            for c in DISPLAY_COLS:
                tree.heading(c, text=c)
                tree.column(c, width=120, anchor='w')
            for c,w in {'검색어':170,'제목':330,'소장기관':170,'거리(km)':75,'자료실':190,'청구기호':150,'상태':90,'반납예정일':100,'상호대차':100,'딸림자료':180}.items():
                tree.column(c,width=w)
            sy=ttk.Scrollbar(frame,orient='vertical',command=tree.yview)
            sx=ttk.Scrollbar(frame,orient='horizontal',command=tree.xview)
            tree.configure(yscrollcommand=sy.set,xscrollcommand=sx.set)
            tree.grid(row=0,column=0,sticky='nsew'); sy.grid(row=0,column=1,sticky='ns'); sx.grid(row=1,column=0,sticky='ew')
            frame.rowconfigure(0,weight=1); frame.columnconfigure(0,weight=1)
            self.trees.append(tree)
        self.rows=[]

    def open_site(self):
        webbrowser.open(BASE)

    def start(self):
        queries = [clean(x) for x in self.input.get('1.0','end').splitlines() if clean(x)]
        if not queries:
            messagebox.showwarning('입력 필요','도서 제목을 한 줄에 하나씩 입력하세요.')
            return
        self.btn.config(state='disabled'); self.export_btn.config(state='disabled')
        self.status.config(text='검색 준비 중...')
        for tree in self.trees:
            fill_tree(tree, [])

        def work():
            try:
                rows = run_queries(queries, lambda m: self.root.after(0, lambda m=m: self.status.config(text=m)))
                self.root.after(0, lambda: self.finish(rows))
            except Exception as e:
                self.root.after(0, lambda e=e: self.fail(e))
        threading.Thread(target=work, daemon=True).start()

    def finish(self, rows):
        self.rows = rows
        avail = [r for r in rows if r.get('상태') == '대출가능']
        other = [r for r in rows if r.get('상태') != '대출가능']
        fill_tree(self.trees[0], avail)
        fill_tree(self.trees[1], other)
        self.status.config(text=f'완료: 전체 {len(rows)}건 / 대출가능 {len(avail)}건')
        self.btn.config(state='normal'); self.export_btn.config(state='normal')
        self.notebook.select(0 if avail else 1)

    def fail(self, e):
        self.status.config(text='오류')
        self.btn.config(state='normal'); self.export_btn.config(state='normal')
        messagebox.showerror('검색 오류', str(e))

    def export(self):
        if not self.rows:
            messagebox.showinfo('안내','먼저 검색을 실행하세요.')
            return
        path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel','*.xlsx')], initialfile='asan_books.xlsx')
        if path:
            try:
                save_xlsx(self.rows, path)
                messagebox.showinfo('저장 완료', path)
            except Exception as e:
                messagebox.showerror('저장 오류', str(e))


if __name__ == '__main__':
    root = tk.Tk()
    App(root)
    root.mainloop()
