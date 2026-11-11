"""crawler.py 수집 결과를 skincare_formulation_DB_300.xlsx 와 같은 형식의 처방 DB로 정리한다.

입력 (crawler.py 를 실행한 폴더)
  formulations_master.csv, ingredients_detail.csv, manufacturing_process.csv, crawl_log.csv

출력
  skincare_formulation_DB_<N>_crawled.xlsx   시트: START_HERE, Formulation_Master, Ingredient_Detail,
                                             Manufacturing_Process, Source_Audit, QA_Check, Crawl_Log
  skincare_formulation_DB_<N>_crawled_*.csv  시트별 CSV (UTF-8 BOM)

사용법
  python3 build_formulation_db.py data/crawl/run300 [--target 300]
"""

from __future__ import annotations

import argparse
import csv
import re
from collections import Counter, defaultdict
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

RIGHTS = "공개 열람; 대량 추출·사내 DB 저장 권한 별도 확인"
SOURCE_SITE = "Personal Care Magazine"

# 제품 유형: 처방명 기준, 위에서부터 먼저 걸리는 규칙 적용
TYPE_RULES = [
    (r"night\s*cream|night\s*cr[eè]me|overnight", "나이트 크림"),
    (r"day\s*cream|day\s*cr[eè]me", "데이 크림"),
    (r"\beye\b|under[- ]eye|contour", "아이 크림"),
    (r"gel[- ]?cream|cream[- ]?gel|gel[- ]?cr[eè]me", "젤 크림"),
    (r"\btoner\b|\btonic\b|lotion tonique|toning", "토너"),
    (r"ampou?le", "앰플"),
    (r"\bserum\b|s[ée]rum", "세럼"),
    (r"essence", "에센스"),
    (r"\bmist\b|spray", "미스트"),
    (r"\bmask\b|masque", "마스크"),
    (r"\bbalm\b|baume", "밤"),
    (r"body\s*(lotion|milk)", "바디 로션"),
    (r"body\s*butter|butter", "바디 버터"),
    (r"sun|spf|uv\b", "선케어"),
    (r"cleans|wash|foam|micellar|make[- ]?up remover|shower|bath", "클렌저"),
    (r"scrub|peel|exfoli", "스크럽/필링"),
    (r"\bstick\b", "스틱"),
    (r"\bfluid\b", "플루이드"),
    (r"emulsion", "에멀전"),
    (r"\bmilk\b|lait", "밀크"),
    (r"lotion", "로션"),
    (r"moistu?ri[sz]", "모이스처라이저"),
    (r"cream|cr[eè]me|creme", "크림"),
    (r"\bgel\b|jelly", "젤"),
    (r"\boil\b|huile", "오일"),
]

NUM_RE = re.compile(r"^\s*(\d+(?:[.,]\d+)?)\s*%?\s*$")


def product_type(name: str) -> str:
    n = name.lower()
    for pat, label in TYPE_RULES:
        if re.search(pat, n):
            return label
    return "기타"


def amount_numeric(raw: str) -> str:
    """'3.0' → 3, '5,68' → 5.68. q.s./up to 100 등은 원문 보존(빈칸)."""
    m = NUM_RE.match(raw or "")
    if not m:
        return ""
    v = float(m.group(1).replace(",", "."))
    return f"{v:g}"


def read_csv(p: Path) -> list[dict]:
    if not p.exists():
        return []
    with open(p, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def txt(v):
    if isinstance(v, str):
        return ILLEGAL_CHARACTERS_RE.sub("", v)[:32000]
    return v


def write_sheet(wb, title: str, header: list[str], rows: list[list], widths: dict[str, int] | None = None):
    ws = wb.create_sheet(title)
    ws.append(header)
    for r in rows:
        ws.append([txt(v) for v in r])
    fill = PatternFill("solid", fgColor="FF315B72")
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFFFF")
        c.fill = fill
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for j, h in enumerate(header, 1):
        ws.column_dimensions[get_column_letter(j)].width = (widths or {}).get(h, 18)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=True)
    ws.freeze_panes = "A2"
    if rows:
        ws.auto_filter.ref = ws.dimensions
    return ws


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("crawl_dir")
    ap.add_argument("--target", type=int, default=300)
    a = ap.parse_args()
    d = Path(a.crawl_dir)

    master = read_csv(d / "formulations_master.csv")
    ingr = read_csv(d / "ingredients_detail.csv")
    proc = {r["formula_url"]: r["preparation"] for r in read_csv(d / "manufacturing_process.csv")}
    logs = read_csv(d / "crawl_log.csv")

    by_url = defaultdict(list)
    for r in ingr:
        by_url[r["formula_url"]].append(r)

    # 성분표까지 파싱된 처방을 수집 순서(카테고리 번갈아 선택된 순서)대로 target 건
    parsed = [m for m in master if by_url.get(m["formula_url"])]
    unparsed = [m for m in master if not by_url.get(m["formula_url"])]
    chosen = parsed[: a.target]
    n = len(chosen)

    fm_rows, id_rows, mp_rows = [], [], []
    totals = {}
    for i, m in enumerate(chosen, 1):
        fid = f"SK{i:03d}"
        url = m["formula_url"]
        rows = by_url[url]
        nums = [amount_numeric(r["amount_raw"]) for r in rows]
        total = sum(float(x) for x in nums if x)
        has_qs = any(not x and r["amount_raw"].strip() for x, r in zip(nums, rows))
        totals[fid] = (total, has_qs)
        fm_rows.append([
            fid, m["formula_name"], product_type(m["formula_name"]), m["supplier_or_publisher"], SOURCE_SITE, url,
            "개별 처방 상세 페이지", "성분/INCI/함량 상세 파싱 완료",
            sum(bool(r["inci"].strip()) for r in rows),
            m["pH"], m["viscosity"], m["appearance"], m["stability"], RIGHTS, m["source_category"],
        ])
        for r, num in zip(rows, nums):
            id_rows.append([fid, m["formula_name"], r["phase"], r["raw_material"], r["inci"],
                            float(num) if num else None, r["amount_raw"], r["supplier"], url])
        mp_rows.append([fid, m["formula_name"], proc.get(url, ""), url])

    # --- QA
    ok_total = sum(1 for t, qs in totals.values() if 99.0 <= t <= 101.0)
    qs_only = sum(1 for t, qs in totals.values() if qs and not 99.0 <= t <= 101.0)
    off = [fid for fid, (t, qs) in totals.items() if not qs and not 99.0 <= t <= 101.0]
    ids = [r[0] for r in fm_rows]
    blocked = [l for l in logs if l["status"] == "BLOCKED"]
    qa = [
        ["총 Formulation_Master 행 수", n, n == a.target],
        ["고유 Formula_ID 수", len(set(ids)), len(set(ids)) == n],
        ["고유 Source_URL 수", len({r[5] for r in fm_rows}), len({r[5] for r in fm_rows}) == n],
        ["Ingredient_Detail 행 수", len(id_rows), len(id_rows) > 0],
        ["성분 없는 Formula_ID", sum(1 for fid in ids if fid not in {r[0] for r in id_rows}), True],
        ["함량 합계 99~101%", ok_total, True],
        ["q.s./up to 100 포함(합계 <100 정상)", qs_only, True],
        ["함량 합계 이상(검토 필요): " + ", ".join(off[:30]), len(off), len(off) == 0],
        ["pH 기재 처방", sum(1 for r in fm_rows if r[9]), True],
        ["제조 방법 기재 처방", sum(1 for r in mp_rows if r[2]), True],
        ["수집했지만 성분표 미파싱(제외)", len(unparsed), True],
        ["접근 제한(BLOCKED) 발생", len(blocked), len(blocked) == 0],
        ["권리 확인 필요 표시", n, True],
    ]

    # --- Source_Audit (카테고리별)
    cat_counts = Counter(r[14] for r in fm_rows)
    cat_urls = {}
    for l in logs:
        m = re.search(r"/formulations-by-category/\d+/[^/?#]+", l["url"])
        if m:
            cat_urls.setdefault(m.group(0), l["url"].split("?")[0])
    from crawler import category_name  # 카테고리 이름 규칙 공유
    audit = []
    for path, url in cat_urls.items():
        name = category_name(url)
        pages = sum(1 for l in logs if path in l["url"] and l["status"] == "OK")
        found = sum(int(re.search(r"(\d+) new", l["message"]).group(1)) for l in logs
                    if path in l["url"] and l["status"] == "OK" and re.search(r"(\d+) new", l["message"]))
        audit.append([f"Personal Care Magazine – {name}", url,
                      f"목록 {pages}페이지, 처방 링크 {found}개 발견",
                      f"{cat_counts.get(name, 0)}건 수록", RIGHTS])

    start = [
        ["목적", f"토너·로션·세럼·앰플·크림 중심 스킨케어 제형 마스터 DB {n}건 (상세 처방 수집본)"],
        ["레코드 수", f"Formulation_Master {n}건, Formula_ID SK001–SK{n:03d}, Ingredient_Detail {len(id_rows)}행"],
        ["수집", f"{date.today().isoformat()} crawler.py 로 Personal Care Magazine 카테고리 목록(페이지 넘김) → 개별 상세 페이지 수집. "
                 "요청 간격 3초, robots.txt 확인, 차단 응답 시 즉시 중단(우회 없음)"],
        ["상세 처방", "모든 레코드가 개별 상세 페이지에서 Phase/상품명/INCI/공급사/함량을 파싱한 처방. 성분표를 읽지 못한 페이지는 제외"],
        ["선택 순서", "카테고리를 번갈아 가며 선택해 한 카테고리 쏠림을 줄임. Source_Category = 처음 발견된 카테고리"],
        ["Product_Type", "처방명 키워드 기반 자동 분류(나이트 크림/세럼/토너/젤 크림 등). 원문 제품 분류가 아님"],
        ["측정 데이터", "pH/점도/외관/안정성은 원문에 'pH', 'Viscosity', 'Appearance', 'Stability' 라벨로 기재된 값만 입력. 공란은 미공개 또는 미확인"],
        ["함량", "Amount_as_Published 원문 보존. 순수 숫자만 Amount_pct_numeric 으로 변환('5,68'→5.68). q.s., up to 100 등은 빈칸"],
        ["INCI 표준화", "원문 표기 그대로(공급사 상품명·복합원료 '(and)' 포함). 표준화하지 않음"],
        ["제조 방법", "Manufacturing_Process 시트에 원문 Preparation 수록"],
        ["권리", "공개 열람·개별 다운로드가 대량 추출·사내 DB 저장·재배포 허가를 의미하지 않음. 회사 내 활용 전 출처별 이용약관 및 서면 허가 확인 필요"],
        ["AI 활용", "pH·점도·외관·안정성 예측 학습은 측정값이 기재된 처방만 분리해 사용"],
    ]

    wb = Workbook()
    wb.remove(wb.active)
    write_sheet(wb, "START_HERE", ["항목", "설명"], start, {"항목": 16, "설명": 120})
    fm_head = ["Formula_ID", "Formula_Name", "Product_Type", "Supplier_or_Publisher", "Source_Site", "Source_URL",
               "Source_Level", "Ingredient_Data_Status", "INCI_Count", "pH", "Viscosity", "Appearance", "Stability",
               "Rights_Status", "Source_Category"]
    write_sheet(wb, "Formulation_Master", fm_head, fm_rows,
                {"Formula_Name": 40, "Supplier_or_Publisher": 26, "Source_URL": 50, "Viscosity": 30, "Appearance": 30,
                 "Rights_Status": 30, "Ingredient_Data_Status": 26})
    id_head = ["Formula_ID", "Formula_Name", "Phase", "Raw_Material_as_Published", "INCI_as_Published",
               "Amount_pct_numeric", "Amount_as_Published", "Supplier", "Source_URL"]
    write_sheet(wb, "Ingredient_Detail", id_head, id_rows,
                {"Formula_Name": 34, "Raw_Material_as_Published": 30, "INCI_as_Published": 50, "Supplier": 26,
                 "Source_URL": 50, "Phase": 8})
    mp_head = ["Formula_ID", "Formula_Name", "Preparation_as_Published", "Source_URL"]
    write_sheet(wb, "Manufacturing_Process", mp_head, mp_rows,
                {"Formula_Name": 34, "Preparation_as_Published": 100, "Source_URL": 50})
    sa_head = ["Source", "URL", "Available_Data", "Used_In_DB", "Rights_and_Usage_Note"]
    write_sheet(wb, "Source_Audit", sa_head, audit, {"Source": 40, "URL": 70, "Available_Data": 30, "Rights_and_Usage_Note": 40})
    write_sheet(wb, "QA_Check", ["검증 항목", "결과", "통과 여부"], qa, {"검증 항목": 60})
    write_sheet(wb, "Crawl_Log", ["URL", "Status", "Message"], [[l["url"], l["status"], l["message"]] for l in logs],
                {"URL": 80, "Message": 50})

    base = d / f"skincare_formulation_DB_{n}_crawled"
    wb.save(base.with_suffix(".xlsx"))
    for ws in wb.worksheets:
        with open(f"{base}_{ws.title}.csv", "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            for row in ws.iter_rows(values_only=True):
                w.writerow(["" if v is None else v for v in row])

    print(f"출력 → {base.with_suffix('.xlsx')} (+ 시트별 CSV)")
    print(f"처방 {n}건 / 성분 {len(id_rows)}행 / 성분표 미파싱 제외 {len(unparsed)}건")
    print("Product_Type:", Counter(r[2] for r in fm_rows).most_common())
    print("Source_Category:", cat_counts.most_common())
    for q in qa:
        print("  QA", q)


if __name__ == "__main__":
    main()
