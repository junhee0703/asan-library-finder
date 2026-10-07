from __future__ import annotations

import re
import threading
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import quote_plus, urljoin

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

BASE_URL = "https://lib.asan.go.kr/dls_le/index.php"
LIBRARY_PRIORITY = {
    "탕정온샘도서관": 1,
    "배방월천도서관": 2,
    "배방도서관": 3,
    "음봉어울샘도서관": 4,
    "꿈샘어린이청소년도서관": 5,
    "중앙도서관": 6,
    "둔포도서관": 7,
    "신창도서관": 8,
    "남산도서관": 9,
}

@dataclass
class Holding:
    query: str
    query_index: int
    title: str = ""
    author: str = ""
    publisher: str = ""
    year: str = ""
    library: str = ""
    location: str = ""
    call_number: str = ""
    status: str = ""
    due_date: str = ""
    reservation: str = ""
    interlibrary: str = ""
    supplement: str = ""
    registration_no: str = ""
    volume: str = ""
    detail_url: str = ""
    marc_url: str = ""
    book_key: str = ""
    priority: int = 999


def clean(value: str | None) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def library_priority(name: str) -> int:
    name = clean(name)
    for library, rank in LIBRARY_PRIORITY.items():
        if library in name:
            return rank
    return 999


def normalize_status(value: str) -> str:
    value = clean(value)
    if "대출가능" in value:
        return "대출가능"
    if "대출중" in value:
        return "대출중"
    if "예약" in value:
        return "예약/기타"
    return value or "확인필요"


def parse_label_line(text: str, first_label: str, second_label: str = "") -> str:
    text = clean(text)
    if first_label not in text:
        return ""
    part = text.split(first_label, 1)[1]
    if second_label and second_label in part:
        part = part.split(second_label, 1)[0]
    part = part.split("|", 1)[0]
    return clean(part)


def soup_get(session: requests.Session, url: str) -> BeautifulSoup:
    r = session.get(url, timeout=30)
    r.raise_for_status()
    r.encoding = r.apparent_encoding or r.encoding or "utf-8"
    return BeautifulSoup(r.text, "html.parser")


def parse_search_card(card, page_url: str, query: str, query_index: int) -> Holding | None:
    link = card.select_one("div.ico.ico-bk a[href*='searchResultDetail']")
    if not link:
        return None
    title = clean(link.get_text(" ", strip=True))
    href = link.get("href", "")
    if not href:
        return None

    detail_url = urljoin(page_url, href)
    key_input = card.select_one("input.listCheck[name='bookKey[]']")
    book_key = clean(key_input.get("value")) if key_input else ""

    author = publisher = year = call_number = ""
    ul = card.select_one("dd > ul")
    lis = ul.find_all("li", recursive=False) if ul else []
    if len(lis) >= 1:
        line1 = clean(lis[0].get_text(" ", strip=True))
        author = parse_label_line(line1, "저자", "발행처")
        publisher = parse_label_line(line1, "발행처")
    if len(lis) >= 2:
        line2 = clean(lis[1].get_text(" ", strip=True))
        year = parse_label_line(line2, "발행년", "청구기호")
        call_number = parse_label_line(line2, "청구기호")

    library = location = ""
    so = card.select_one("li.so")
    if so:
        lib_span = so.select_one("span.fb.blue")
        loc_span = so.select_one("span.fb.yellow")
        library = clean(lib_span.get_text(" ", strip=True)) if lib_span else ""
        location = clean(loc_span.get_text(" ", strip=True)) if loc_span else ""

    status_node = card.select_one("ol strong")
    status = normalize_status(status_node.get_text(" ", strip=True)) if status_node else ""
    interlibrary = ""
    for a in card.select("ol a"):
        txt = clean(a.get_text(" ", strip=True))
        if "상호대차" in txt:
            interlibrary = txt
            break

    return Holding(query=query, query_index=query_index, title=title,
                   author=author, publisher=publisher, year=year,
                   library=library, location=location, call_number=call_number,
                   status=status, interlibrary=interlibrary,
                   detail_url=detail_url, book_key=book_key,
                   priority=library_priority(library))


def total_count(soup: BeautifulSoup) -> int:
    h3 = soup.select_one("h3")
    text = clean(h3.get_text(" ", strip=True)) if h3 else ""
    m = re.search(r"총\s*([\d,]+)\s*권", text)
    if not m:
        m = re.search(r"총\s*([\d,]+)\s*권\(개\)", clean(soup.get_text(" ", strip=True)))
    return int(m.group(1).replace(",", "")) if m else 0


def search_url(query: str, offset: int = 1, list_num: int = 50) -> str:
    q = quote_plus(query)
    return (f"{BASE_URL}?preWord={q}&deSearch=2&item=total&word={q}"
            f"&act=dataSearchForm&listNum={list_num}&offset={offset}")


def search_one(session: requests.Session, query: str, query_index: int, progress_callback=None) -> list[Holding]:
    rows, seen = [], set()
    list_num, offset = 50, 1
    while True:
        url = search_url(query, offset, list_num)
        if progress_callback:
            progress_callback(f"검색 중: {query}")
        soup = soup_get(session, url)
        cards = [x for x in soup.select("div.list") if x.select_one("input.listCheck[name='bookKey[]']")]
        for card in cards:
            parsed = parse_search_card(card, url, query, query_index)
            if parsed and parsed.detail_url not in seen:
                seen.add(parsed.detail_url)
                rows.append(parsed)
        total = total_count(soup)
        if total <= 0 or len(rows) >= total or len(cards) < list_num or offset > 5000:
            break
        offset += list_num
    return rows


def parse_detail_row(row, holding: Holding) -> None:
    def cell(name: str) -> str:
        el = row.select_one(f"td[data-th='{name}']")
        return clean(el.get_text(" ", strip=True)) if el else ""
    reg = row.select_one("td[data-th='구분']")
    if reg:
        raw = clean(reg.get_text(" ", strip=True))
        if raw:
            holding.registration_no = raw.split()[0]
    holding.volume = cell("낱권정보")
    room_cell = row.select_one("td[data-th='자료실 / 청구기호']")
    if room_cell:
        strong = room_cell.select_one("strong")
        call = room_cell.select_one("a.print")
        if strong:
            holding.location = clean(strong.get_text(" ", strip=True))
        if call:
            holding.call_number = clean(call.get_text(" ", strip=True))
    holding.status = normalize_status(cell("자료상태"))
    holding.due_date = cell("반납예정일")
    holding.reservation = cell("예약")
    holding.interlibrary = cell("상호대차") or holding.interlibrary


def extract_supplement_from_text(text: str) -> str:
    text = clean(text)
    if not text:
        return ""
    patterns = [
        r"딸림자료\s*[:：]?\s*([^\n|]+)",
        r"부록\s*[:：]?\s*([^\n|]+)",
        r"부록자료\s*[:：]?\s*([^\n|]+)",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.I)
        if m:
            value = clean(m.group(1))
            if value and value not in {"-", "없음", "없다", "무"}:
                return value
            return "없음"

    # MARC 300 등에서 흔히 보이는 물리적 구성 표기: "+ CD-ROM 1매", "+ 별책 1책" 등
    m = re.search(r"\+\s*([^\n|]{1,100})", text, re.I)
    if m:
        value = clean(m.group(1))
        if re.search(r"CD|DVD|별책|부록|책자|매|개", value, re.I):
            return value
    m = re.search(r"(?:CD-ROM|CD|DVD)\s*\d*\s*(?:매|장)?", text, re.I)
    if m:
        return m.group(0)
    return ""


def enrich_from_detail(session: requests.Session, holding: Holding) -> None:
    try:
        soup = soup_get(session, holding.detail_url)
        rows = soup.select("table.tstyle.responsive tbody tr")
        matched = False
        for row in rows:
            cb = row.select_one("input[name='bookKey[]']")
            if holding.book_key and cb and clean(cb.get("value")) == holding.book_key:
                parse_detail_row(row, holding)
                matched = True
                break
        if not matched and len(rows) == 1:
            parse_detail_row(rows[0], holding)

        marc = soup.find("a", string=lambda s: s and "marc 보기" in clean(s))
        if marc:
            holding.marc_url = urljoin(holding.detail_url, marc.get("href", ""))
            if holding.marc_url:
                try:
                    marc_soup = soup_get(session, holding.marc_url)
                    marc_text = marc_soup.get_text(" ", strip=True)
                    holding.supplement = extract_supplement_from_text(marc_text)
                except Exception:
                    pass

        if not holding.supplement:
            holding.supplement = extract_supplement_from_text(soup.get_text(" ", strip=True)) or "없음"
        holding.priority = library_priority(holding.library)
    except Exception:
        if not holding.supplement:
            holding.supplement = "확인불가"


def sort_rows(rows: list[Holding]) -> list[Holding]:
    return sorted(rows, key=lambda r: (r.query_index, 0 if r.status == "대출가능" else 1,
                                       r.priority, r.title.lower(), r.library))


def save_excel(path: Path, rows: list[Holding]) -> None:
    wb = Workbook()
    headers = ["검색어", "제목", "저자", "출판사", "발행년", "소장기관", "자료실",
               "청구기호", "상태", "반납예정일", "상호대차", "딸림자료", "등록번호", "권차"]
    ordered = sort_rows(rows)
    ws = wb.active
    ws.title = "전체"
    ws.append(headers)
    for r in ordered:
        ws.append(row_values(r))
    for title, subset in (("대출가능", [r for r in ordered if r.status == "대출가능"]),
                          ("대출중_기타", [r for r in ordered if r.status != "대출가능"])):
        sh = wb.create_sheet(title)
        sh.append(headers)
        for r in subset:
            sh.append(row_values(r))
    for sh in wb.worksheets:
        sh.freeze_panes = "A2"
        sh.auto_filter.ref = sh.dimensions
        for i, w in enumerate([22,46,28,22,10,24,28,24,14,16,22,18,18,12], 1):
            sh.column_dimensions[chr(64+i)].width = w
    wb.save(path)


def row_values(r: Holding):
    return [r.query, r.title, r.author, r.publisher, r.year, r.library, r.location,
            r.call_number, r.status, r.due_date, r.interlibrary, r.supplement,
            r.registration_no, r.volume]


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("아산시립도서관 책 찾기")
        root.geometry("1500x850")
        root.minsize(1100, 650)
        self.rows = []
        tk.Label(root, text="도서 제목 — 한 줄에 하나씩", font=("Malgun Gothic", 11, "bold")).pack(anchor="w", padx=12, pady=(10,3))
        self.input = tk.Text(root, height=5, font=("Malgun Gothic", 11)); self.input.pack(fill="x", padx=12)
        self.input.insert("1.0", "Princess in Black\nJigsaw Jones Mystery\nMarvin Redpost\nThe Zack Files\nJudy Moody")
        bar = tk.Frame(root); bar.pack(fill="x", padx=12, pady=8)
        self.search_btn = tk.Button(bar, text="검색 시작", width=12, command=self.start_search); self.search_btn.pack(side="left")
        tk.Button(bar, text="엑셀 저장", width=12, command=self.export_excel).pack(side="left", padx=8)
        self.status_var = tk.StringVar(value="대기 중")
        tk.Label(bar, textvariable=self.status_var, font=("Malgun Gothic", 10, "bold")).pack(side="left", padx=18)
        self.notebook = ttk.Notebook(root); self.notebook.pack(fill="both", expand=True, padx=12, pady=(0,12))
        self.available_tree = self.make_tree(self.notebook)
        self.borrowed_tree = self.make_tree(self.notebook)
        self.notebook.add(self.available_tree, text="대출가능")
        self.notebook.add(self.borrowed_tree, text="대출중 / 기타")

    def make_tree(self, parent):
        cols = ("검색어","제목","소장기관","자료실","청구기호","상태","반납예정일","상호대차","딸림자료")
        frame = ttk.Frame(parent); frame.rowconfigure(0, weight=1); frame.columnconfigure(0, weight=1)
        tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="browse")
        widths = {"검색어":170,"제목":430,"소장기관":190,"자료실":230,"청구기호":180,"상태":90,"반납예정일":110,"상호대차":150,"딸림자료":180}
        for c in cols:
            tree.heading(c, text=c); tree.column(c, width=widths[c], anchor="w", stretch=False)
        y = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        x = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=y.set, xscrollcommand=x.set)
        tree.grid(row=0,column=0,sticky="nsew"); y.grid(row=0,column=1,sticky="ns"); x.grid(row=1,column=0,sticky="ew")
        frame.grid(row=0,column=0,sticky="nsew")
        parent.rowconfigure(0, weight=1); parent.columnconfigure(0, weight=1)
        # 마우스 휠 + 키보드 방향키로 상하/좌우 이동을 확실하게 지원
        tree.bind("<MouseWheel>", lambda e, t=tree: self._wheel_y(e, t))
        tree.bind("<Shift-MouseWheel>", lambda e, t=tree: self._wheel_x(e, t))
        tree.bind("<Button-4>", lambda e, t=tree: self._wheel_y_linux(e, t, -1))
        tree.bind("<Button-5>", lambda e, t=tree: self._wheel_y_linux(e, t, 1))
        tree.bind("<Left>", lambda e, t=tree: self._key_x(e, t, -1))
        tree.bind("<Right>", lambda e, t=tree: self._key_x(e, t, 1))
        tree.bind("<Up>", lambda e, t=tree: self._key_y(e, t, -1))
        tree.bind("<Down>", lambda e, t=tree: self._key_y(e, t, 1))
        tree.bind("<Prior>", lambda e, t=tree: self._page_y(e, t, -1))
        tree.bind("<Next>", lambda e, t=tree: self._page_y(e, t, 1))
        return frame

    def _wheel_y(self, e, t):
        t.yview_scroll(-int(e.delta/120), "units"); return "break"
    def _wheel_x(self, e, t):
        t.xview_scroll(-int(e.delta/120), "units"); return "break"
    def _wheel_y_linux(self, e, t, n):
        t.yview_scroll(n, "units"); return "break"
    def _key_x(self, e, t, n):
        t.xview_scroll(n, "units"); return "break"
    def _key_y(self, e, t, n):
        t.yview_scroll(n, "units"); return "break"
    def _page_y(self, e, t, n):
        t.yview_scroll(n, "pages"); return "break"

    def start_search(self):
        if getattr(self, "_searching", False): return
        queries = [clean(x) for x in self.input.get("1.0", "end").splitlines() if clean(x)]
        if not queries:
            messagebox.showwarning("검색", "도서 제목을 한 줄에 하나씩 입력해 주세요."); return
        self._searching = True; self.search_btn.config(state="disabled"); self.rows=[]; self.clear_trees()
        threading.Thread(target=self.worker, args=(queries,), daemon=True).start()

    def worker(self, queries):
        try:
            with requests.Session() as session:
                session.headers.update({"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/154 Safari/537.36","Accept-Language":"ko-KR,ko;q=0.9,en;q=0.8"})
                all_rows=[]
                for idx, query in enumerate(queries):
                    try:
                        rows=search_one(session,query,idx,lambda msg:self.root.after(0,self.status_var.set,msg))
                        for n,row in enumerate(rows):
                            self.root.after(0,self.status_var.set,f"{query}: 상세정보 확인 {n+1}/{len(rows)}")
                            enrich_from_detail(session,row)
                        all_rows.extend(rows)
                    except Exception as e:
                        all_rows.append(Holding(query=query,query_index=idx,title="검색 중 오류",status=f"오류: {str(e)[:80]}",supplement="확인불가"))
                self.rows=sort_rows(all_rows); self.root.after(0,self.refresh_trees)
        except Exception as e:
            self.root.after(0,lambda:messagebox.showerror("검색 오류",str(e)))
        finally:
            self.root.after(0,self.finish_search)

    def finish_search(self):
        self._searching=False; self.search_btn.config(state="normal")
        available=sum(r.status=="대출가능" for r in self.rows)
        self.status_var.set(f"완료: 전체 {len(self.rows)}건 / 대출가능 {available}건")
    def clear_trees(self):
        for tree in (self.available_tree,self.borrowed_tree):
            for item in tree.get_children(): tree.delete(item)
    def refresh_trees(self):
        self.clear_trees()
        for r in self.rows:
            values=(r.query,r.title,r.library,r.location,r.call_number,r.status,r.due_date,r.interlibrary,r.supplement)
            tree=self.available_tree if r.status=="대출가능" else self.borrowed_tree
            tree.insert("","end",values=values)
    def export_excel(self):
        if not self.rows:
            messagebox.showinfo("엑셀 저장","먼저 검색을 실행해 주세요."); return
        path=filedialog.asksaveasfilename(title="검색 결과 저장",defaultextension=".xlsx",filetypes=[("Excel 파일","*.xlsx")],initialfile="아산도서관_검색결과.xlsx")
        if not path:return
        try:
            save_excel(Path(path),self.rows); messagebox.showinfo("엑셀 저장",f"저장했습니다.\n{path}")
        except Exception as e: messagebox.showerror("엑셀 저장 오류",str(e))


def main():
    root=tk.Tk(); App(root); root.mainloop()
if __name__=="__main__": main()
