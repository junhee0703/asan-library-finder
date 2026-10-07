from __future__ import annotations

import re
import sys
import threading
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import quote_plus, urljoin

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from openpyxl import Workbook
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError


BASE_URL = "https://lib.asan.go.kr/dls_le/index.php"

# 탕정면을 기준으로 한 내부 표시 순서.
# 화면에는 거리(km)를 표시하지 않고, 이 순서만 검색 결과 정렬에 사용합니다.
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
    # 예: "저자 [by] Shannon Hale | 발행처 Walker Books"
    if first_label not in text:
        return ""
    part = text.split(first_label, 1)[1]
    if second_label and second_label in part:
        part = part.split(second_label, 1)[0]
    part = part.split("|", 1)[0]
    return clean(part)


def parse_search_card(card, page_url: str, query: str, query_index: int) -> Holding | None:
    link = card.locator("div.ico.ico-bk a[href*='searchResultDetail']").first
    if link.count() == 0:
        return None

    title = clean(link.inner_text())
    href = link.get_attribute("href") or ""
    if not href:
        return None

    detail_url = urljoin(page_url, href)

    book_key = ""
    key_input = card.locator("input.listCheck[name='bookKey[]']").first
    if key_input.count():
        book_key = clean(key_input.get_attribute("value"))

    ul = card.locator("dd ul").first
    lis = ul.locator(":scope > li") if ul.count() else None

    author = publisher = year = call_number = ""
    if lis is not None:
        n = lis.count()
        if n >= 1:
            line1 = clean(lis.nth(0).inner_text())
            author = parse_label_line(line1, "저자", "발행처")
            publisher = parse_label_line(line1, "발행처")
        if n >= 2:
            line2 = clean(lis.nth(1).inner_text())
            year = parse_label_line(line2, "발행년", "청구기호")
            call_number = parse_label_line(line2, "청구기호")

    library = ""
    location = ""
    so = card.locator("li.so").first
    if so.count():
        lib_span = so.locator("span.fb.blue").first
        loc_span = so.locator("span.fb.yellow").first
        if lib_span.count():
            library = clean(lib_span.inner_text())
        if loc_span.count():
            location = clean(loc_span.inner_text())

    status = ""
    status_node = card.locator("ol strong").first
    if status_node.count():
        status = normalize_status(status_node.inner_text())

    interlibrary = ""
    for i in range(card.locator("ol a").count()):
        a = card.locator("ol a").nth(i)
        txt = clean(a.inner_text())
        if "상호대차" in txt:
            interlibrary = txt
            break
    if not interlibrary:
        ol_text = clean(card.locator("ol").inner_text()) if card.locator("ol").count() else ""
        if "상호대차" in ol_text:
            interlibrary = ol_text

    return Holding(
        query=query,
        query_index=query_index,
        title=title,
        author=author,
        publisher=publisher,
        year=year,
        library=library,
        location=location,
        call_number=call_number,
        status=status,
        interlibrary=interlibrary,
        detail_url=detail_url,
        book_key=book_key,
        priority=library_priority(library),
    )


def total_count(page) -> int:
    text = clean(page.locator("h3").first.inner_text()) if page.locator("h3").count() else ""
    m = re.search(r"총\s*([\d,]+)\s*권", text)
    if not m:
        m = re.search(r"총\s*([\d,]+)\s*권\(개\)", page.locator("body").inner_text())
    return int(m.group(1).replace(",", "")) if m else 0


def search_url(query: str, offset: int = 1, list_num: int = 50) -> str:
    q = quote_plus(query)
    params = (
        f"preWord={q}&deSearch=2&item=total&word={q}"
        f"&act=dataSearchForm&listNum={list_num}&offset={offset}"
    )
    return f"{BASE_URL}?{params}"


def search_one(page, query: str, query_index: int, progress_callback=None) -> list[Holding]:
    rows: list[Holding] = []
    seen = set()
    list_num = 50
    offset = 1

    while True:
        url = search_url(query, offset=offset, list_num=list_num)
        if progress_callback:
            progress_callback(f"검색 중: {query} (offset {offset})")

        page.goto(url, wait_until="domcontentloaded", timeout=45000)
        page.wait_for_timeout(700)

        cards = page.locator("div.list").filter(
            has=page.locator("input.listCheck[name='bookKey[]']")
        )
        count = cards.count()

        if count == 0:
            # 혹시 locator의 has 필터가 브라우저 버전에서 불안정하면 직접 검사합니다.
            all_lists = page.locator("div.list")
            count = all_lists.count()
            for i in range(count):
                card = all_lists.nth(i)
                if card.locator("input.listCheck[name='bookKey[]']").count() == 0:
                    continue
                parsed = parse_search_card(card, page.url, query, query_index)
                if parsed and parsed.detail_url not in seen:
                    seen.add(parsed.detail_url)
                    rows.append(parsed)
        else:
            for i in range(count):
                parsed = parse_search_card(cards.nth(i), page.url, query, query_index)
                if parsed and parsed.detail_url not in seen:
                    seen.add(parsed.detail_url)
                    rows.append(parsed)

        total = total_count(page)
        if total <= 0 or len(rows) >= total or count < list_num:
            break
        offset += list_num

        # 무한 반복 방지
        if offset > 5000:
            break

    return rows


def parse_detail_row(row, holding: Holding) -> None:
    def cell(data_th: str) -> str:
        loc = row.locator(f"td[data-th='{data_th}']").first
        return clean(loc.inner_text()) if loc.count() else ""

    reg_cell = row.locator("td[data-th='구분']").first
    if reg_cell.count():
        raw = clean(reg_cell.inner_text())
        if raw:
            holding.registration_no = raw.split()[0]

    holding.volume = cell("낱권정보")

    room_cell = row.locator("td[data-th='자료실 / 청구기호']").first
    if room_cell.count():
        strong = room_cell.locator("strong").first
        call = room_cell.locator("a.print").first
        if strong.count():
            holding.location = clean(strong.inner_text())
        if call.count():
            holding.call_number = clean(call.inner_text())

    holding.status = normalize_status(cell("자료상태"))
    holding.due_date = cell("반납예정일")
    holding.reservation = cell("예약")
    holding.interlibrary = cell("상호대차") or holding.interlibrary


def enrich_from_detail(detail_page, holding: Holding) -> None:
    try:
        detail_page.goto(holding.detail_url, wait_until="domcontentloaded", timeout=45000)
        detail_page.wait_for_timeout(300)

        rows = detail_page.locator("table.tstyle.responsive tbody tr")
        matched = False

        for i in range(rows.count()):
            row = rows.nth(i)
            if holding.book_key:
                cb = row.locator("input[name='bookKey[]']").first
                if cb.count() and clean(cb.get_attribute("value")) == holding.book_key:
                    parse_detail_row(row, holding)
                    matched = True
                    break

        if not matched and rows.count() == 1:
            parse_detail_row(rows.nth(0), holding)

        # 딸림자료 표기가 있는 상세페이지에서만 추출합니다.
        body = detail_page.locator("body").inner_text()
        if "딸림자료" in body:
            m = re.search(r"딸림자료\s*[:：]?\s*([^\n]+)", body)
            if m:
                holding.supplement = clean(m.group(1))

        if not holding.supplement:
            holding.supplement = "-"

        holding.priority = library_priority(holding.library)
    except Exception:
        if not holding.supplement:
            holding.supplement = "-"


def sort_rows(rows: list[Holding]) -> list[Holding]:
    # 사용자가 입력한 도서 순서를 가장 먼저 유지하고,
    # 각 도서 안에서는 대출가능 → 가까운 도서관 순으로 표시합니다.
    return sorted(
        rows,
        key=lambda r: (
            r.query_index,
            0 if r.status == "대출가능" else 1,
            r.priority,
            r.title.lower(),
            r.library,
        ),
    )


def save_excel(path: Path, rows: list[Holding]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "전체"

    headers = [
        "검색어", "제목", "저자", "출판사", "발행년",
        "소장기관", "자료실", "청구기호", "상태",
        "반납예정일", "상호대차", "딸림자료", "등록번호", "권차", "상세URL"
    ]
    ws.append(headers)

    ordered = sort_rows(rows)
    for r in ordered:
        ws.append([
            r.query, r.title, r.author, r.publisher, r.year,
            r.library, r.location, r.call_number, r.status,
            r.due_date, r.interlibrary, r.supplement,
            r.registration_no, r.volume, r.detail_url,
        ])

    for title, subset in (
        ("대출가능", [r for r in ordered if r.status == "대출가능"]),
        ("대출중_기타", [r for r in ordered if r.status != "대출가능"]),
    ):
        sh = wb.create_sheet(title)
        sh.append(headers)
        for r in subset:
            sh.append([
                r.query, r.title, r.author, r.publisher, r.year,
                r.library, r.location, r.call_number, r.status,
                r.due_date, r.interlibrary, r.supplement,
                r.registration_no, r.volume, r.detail_url,
            ])
        sh.freeze_panes = "A2"
        sh.auto_filter.ref = sh.dimensions

    for sh in wb.worksheets:
        sh.freeze_panes = "A2"
        sh.auto_filter.ref = sh.dimensions
        widths = [22, 46, 28, 22, 10, 24, 28, 24, 14, 16, 22, 14, 18, 12, 65]
        for idx, width in enumerate(widths, 1):
            sh.column_dimensions[chr(64 + idx)].width = width

    wb.save(path)


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("아산시립도서관 책 찾기")
        self.root.geometry("1500x850")
        self.root.minsize(1100, 650)

        self.rows: list[Holding] = []
        self.last_search_url = BASE_URL

        tk.Label(root, text="도서 제목 — 한 줄에 하나씩", font=("Malgun Gothic", 11, "bold")).pack(
            anchor="w", padx=12, pady=(10, 3)
        )

        self.input = tk.Text(root, height=5, font=("Malgun Gothic", 11))
        self.input.pack(fill="x", padx=12)
        self.input.insert("1.0", "Princess in Black\nJigsaw Jones Mystery\nMarvin Redpost\nThe Zack Files\nJudy Moody")

        bar = tk.Frame(root)
        bar.pack(fill="x", padx=12, pady=8)

        self.search_btn = tk.Button(bar, text="검색 시작", width=12, command=self.start_search)
        self.search_btn.pack(side="left")

        tk.Button(bar, text="엑셀 저장", width=12, command=self.export_excel).pack(side="left", padx=8)
        tk.Button(bar, text="검색 페이지 열기", width=14, command=self.open_search_page).pack(side="left")

        self.status_var = tk.StringVar(value="대기 중")
        tk.Label(bar, textvariable=self.status_var, font=("Malgun Gothic", 10, "bold")).pack(
            side="left", padx=18
        )

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        self.available_tree = self.make_tree(self.notebook)
        self.borrowed_tree = self.make_tree(self.notebook)
        self.notebook.add(self.available_tree, text="대출가능")
        self.notebook.add(self.borrowed_tree, text="대출중 / 기타")

    def make_tree(self, parent):
        cols = (
            "검색어", "제목", "소장기관", "자료실",
            "청구기호", "상태", "반납예정일", "상호대차", "딸림자료"
        )
        tree = ttk.Treeview(parent, columns=cols, show="headings")
        widths = {
            "검색어": 170, "제목": 430, "소장기관": 190, "자료실": 230,
            "청구기호": 180, "상태": 90, "반납예정일": 110,
            "상호대차": 150, "딸림자료": 110
        }
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=widths[c], anchor="w")
        y = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        x = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=y.set, xscrollcommand=x.set)
        tree.grid(row=0, column=0, sticky="nsew")
        y.grid(row=0, column=1, sticky="ns")
        x.grid(row=1, column=0, sticky="ew")
        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)
        return tree

    def start_search(self):
        if getattr(self, "_searching", False):
            return
        queries = [clean(x) for x in self.input.get("1.0", "end").splitlines() if clean(x)]
        if not queries:
            messagebox.showwarning("검색", "도서 제목을 한 줄에 하나씩 입력해 주세요.")
            return

        self._searching = True
        self.search_btn.config(state="disabled")
        self.rows = []
        self.clear_trees()
        threading.Thread(target=self.worker, args=(queries,), daemon=True).start()

    def worker(self, queries):
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=True)
                context = browser.new_context(locale="ko-KR", viewport={"width": 1440, "height": 1000})
                search_page = context.new_page()
                detail_page = context.new_page()

                all_rows = []
                for idx, query in enumerate(queries):
                    try:
                        rows = search_one(
                            search_page, query, idx,
                            progress_callback=lambda msg: self.root.after(0, self.status_var.set, msg)
                        )
                        # 상세 페이지는 검색결과에 실제로 존재하는 각 소장정보의 반납일 등을 보강할 때만 사용
                        for n, row in enumerate(rows):
                            if self.root.winfo_exists():
                                self.root.after(
                                    0, self.status_var.set,
                                    f"{query}: 상세정보 확인 {n + 1}/{len(rows)}"
                                )
                            enrich_from_detail(detail_page, row)
                        all_rows.extend(rows)
                    except Exception as e:
                        all_rows.append(Holding(
                            query=query,
                            query_index=idx,
                            title="검색 중 오류",
                            status=f"오류: {str(e)[:80]}",
                            supplement="-",
                        ))

                browser.close()
                self.rows = sort_rows(all_rows)
                self.root.after(0, self.refresh_trees)
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("검색 오류", str(e)))
        finally:
            self.root.after(0, self.finish_search)

    def finish_search(self):
        self._searching = False
        self.search_btn.config(state="normal")
        available = sum(1 for r in self.rows if r.status == "대출가능")
        self.status_var.set(f"완료: 전체 {len(self.rows)}건 / 대출가능 {available}건")

    def clear_trees(self):
        for tree in (self.available_tree, self.borrowed_tree):
            for item in tree.get_children():
                tree.delete(item)

    def refresh_trees(self):
        self.clear_trees()
        for r in self.rows:
            values = (
                r.query, r.title, r.library, r.location,
                r.call_number, r.status, r.due_date,
                r.interlibrary, r.supplement,
            )
            tree = self.available_tree if r.status == "대출가능" else self.borrowed_tree
            tree.insert("", "end", values=values)

    def export_excel(self):
        if not self.rows:
            messagebox.showinfo("엑셀 저장", "먼저 검색을 실행해 주세요.")
            return
        path = filedialog.asksaveasfilename(
            title="검색 결과 저장",
            defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx")],
            initialfile="아산도서관_검색결과.xlsx",
        )
        if not path:
            return
        try:
            save_excel(Path(path), self.rows)
            messagebox.showinfo("엑셀 저장", f"저장했습니다.\n{path}")
        except Exception as e:
            messagebox.showerror("엑셀 저장 오류", str(e))

    def open_search_page(self):
        import webbrowser
        webbrowser.open(self.last_search_url)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
