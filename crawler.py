"""
Personal Care Magazine formulation collector
- Collects public formulation detail pages linked from category/archive pages.
- Extracts only fields explicitly present in the page.
- Does not bypass login, CAPTCHA, Cloudflare, paywalls, or access restrictions.
- Check the site's Terms, robots.txt, and obtain permission for bulk collection /
  internal database or AI/ML use before running at scale.
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
import re
import time
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Optional
from urllib.parse import urljoin, urlparse
from urllib.robotparser import RobotFileParser

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE = "https://www.personalcaremagazine.com"
USER_AGENT = "CosmeticFormulationResearchBot/1.0"
DEFAULT_CATEGORIES = [
    f"{BASE}/formulations-by-category/33/skin-treatment",
    f"{BASE}/formulations-by-category/15/facial-care-cleansing",
    f"{BASE}/formulations-by-category/30/anti-ageing",
    f"{BASE}/formulations-by-category/31/skin-brightening",
    f"{BASE}/formulations-by-category/22/cleansing-moisturising",
    f"{BASE}/formulations-by-category/14/body-care-cleansing",
]
CATEGORY_NAMES = {
    "33": "Skin Treatment", "15": "Facial Care/Cleansing", "30": "Anti-Ageing",
    "31": "Skin Brightening", "22": "Cleansing/Moisturising", "14": "Body Care/Cleansing",
}

def category_name(url: str) -> str:
    m = re.search(r"/formulations-by-category/(\d+)/([^/?#]+)", url)
    if not m:
        return ""
    return CATEGORY_NAMES.get(m.group(1), m.group(2).replace("-", " ").title())

@dataclass
class Ingredient:
    formula_url: str
    formula_name: str
    phase: str
    raw_material: str
    inci: str
    supplier: str
    amount_raw: str
    amount_numeric: Optional[float]

@dataclass
class Formula:
    formula_url: str
    formula_name: str
    supplier_or_publisher: str
    source_category: str
    product_type: str
    pH: str
    viscosity: str
    appearance: str
    stability: str
    ingredient_count: int
    parse_status: str

def clean(s: str) -> str:
    return re.sub(r"\s+", " ", s or "").strip()

def amount_to_number(value: str) -> Optional[float]:
    """Return a numeric value only for a plain numeric amount; preserve all raw text."""
    s = clean(value).replace(",", "")
    if not s:
        return None
    s = s.replace("%", "").strip()
    if re.fullmatch(r"\d+(?:\.\d+)?", s):
        return float(s)
    return None

def allowed_by_robots(session: requests.Session, url: str) -> bool:
    parsed = urlparse(url)
    robots_url = f"{parsed.scheme}://{parsed.netloc}/robots.txt"
    rp = RobotFileParser()
    try:
        response = session.get(robots_url, timeout=20)
        response.raise_for_status()
        rp.parse(response.text.splitlines())
    except requests.RequestException as exc:
        raise RuntimeError(
            f"robots.txt를 확인할 수 없어 안전을 위해 중단합니다: {robots_url} ({exc})"
        )
    return rp.can_fetch(USER_AGENT, url)

def make_session() -> requests.Session:
    s = requests.Session()
    s.headers.update({
        "User-Agent": USER_AGENT,
        "Accept": "text/html,application/xhtml+xml",
        "Accept-Language": "en-US,en;q=0.9",
    })
    return s

def get_soup(session: requests.Session, url: str, delay: float) -> BeautifulSoup:
    if urlparse(url).netloc != urlparse(BASE).netloc:
        raise ValueError(f"허용되지 않은 도메인: {url}")
    if not allowed_by_robots(session, url):
        raise PermissionError(f"robots.txt에서 접근을 허용하지 않는 URL: {url}")

    time.sleep(delay)
    response = session.get(url, timeout=30)
    if response.status_code in (401, 403, 429):
        raise PermissionError(
            f"접근 제한 응답({response.status_code}). 우회하지 않고 수집을 중단합니다: {url}"
        )
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "html.parser")
    if is_challenge_page(response.text, soup):
        raise PermissionError(f"자동화 차단/사람 확인 페이지 감지. 수집 중단: {url}")
    return soup

CHALLENGE_MARKERS = (
    "verify you are human", "are you a robot", "cf-challenge", "cf_chl_", "challenge-platform",
    "just a moment...", "attention required!", "access denied", "please complete the security check",
)

def is_challenge_page(html: str, soup: BeautifulSoup) -> bool:
    """실제 차단/사람 확인 페이지인지 판단.
    회원가입·뉴스레터 폼에 들어 있는 reCAPTCHA 입력칸(정상 페이지에도 있음)은 차단으로 보지 않는다.
    차단 페이지로 판단되면 우회하지 않고 수집을 중단한다."""
    low = html.lower()
    if any(m in low for m in CHALLENGE_MARKERS):
        return True
    title = (soup.title.get_text(" ", strip=True).lower() if soup.title else "")
    if "captcha" in title:
        return True
    # 본문 콘텐츠 없이 캡차만 있는 페이지
    if "captcha" in low and not soup.find("h1") and len(soup.get_text(" ", strip=True)) < 1500:
        return True
    return False

def collect_detail_links(soup: BeautifulSoup, page_url: str) -> list[str]:
    links = set()
    for a in soup.select('a[href]'):
        href = urljoin(page_url, a.get("href", ""))
        parsed = urlparse(href)
        if parsed.netloc == urlparse(BASE).netloc and re.search(
            r"/formulation-details/\d+/", parsed.path
        ):
            links.add(href.split("#")[0])
    return sorted(links)

def infer_product_type(title: str, category: str = "") -> str:
    t = f"{title} {category}".lower()
    rules = [
        ("toner", "토너"), ("lotion", "로션"), ("serum", "세럼"),
        ("ampoule", "앰플"), ("cream", "크림"), ("moistur", "크림/보습"),
        ("eye", "아이케어"), ("gel", "젤"), ("essence", "에센스"),
    ]
    found = [label for key, label in rules if key in t]
    return "/".join(dict.fromkeys(found)) if found else "미분류"

def find_value_near_label(text: str, labels: list[str]) -> str:
    for label in labels:
        m = re.search(
            # 단어 경계: 'pH' 가 'Phase' 안의 'ph' 에 걸리지 않도록
            rf"(?<![A-Za-z]){re.escape(label)}(?![A-Za-z])\s*[:\-]?\s*([^\n\r]+)",
            text, flags=re.IGNORECASE
        )
        if m:
            return clean(m.group(1))[:300]
    return ""

def extract_quality(text: str) -> tuple[str, str, str, str]:
    # Conservative: capture only explicit label/value pairs.
    # 'pH 5.5 – 6.5', 'pH: 5,68', 'pH value (23°C) : 5.40' → 숫자(범위)만
    m = re.search(r"(?<![A-Za-z])pH(?:\s*value)?\s*(?:\([^)]*\))?\s*[:\-=]?\s*"
                  r"(\d+(?:[.,]\d+)?(?:\s*[–~-]\s*\d+(?:[.,]\d+)?)?)", text)
    ph = m.group(1) if m else find_value_near_label(text, ["pH"])
    viscosity = find_value_near_label(text, ["Viscosity"])
    appearance = find_value_near_label(text, ["Appearance", "Aspect"])
    stability = find_value_near_label(text, ["Stability", "Stable"])
    return ph, viscosity, appearance, stability

def extract_supplier(soup: BeautifulSoup) -> str:
    # Supplier is commonly the first heading/link immediately below the title.
    h1 = soup.find(["h1"])
    if h1:
        nxt = h1.find_next(["h2", "h3", "h4", "h5", "a"])
        if nxt:
            candidate = clean(nxt.get_text(" ", strip=True))
            if candidate and candidate.lower() not in {"preparation", "phase"}:
                return candidate
    # fallback to visible links in the detail header area
    for a in soup.select("a[href]"):
        txt = clean(a.get_text(" ", strip=True))
        if txt and "supplier" not in txt.lower() and len(txt) > 3:
            if "/formulation-details/" not in a.get("href", ""):
                return txt
    return ""

def parse_ingredient_table(soup: BeautifulSoup, formula_url: str, formula_name: str):
    ingredients: list[Ingredient] = []
    raw_tables = []
    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        if not rows:
            continue
        header_idx = None
        headers = []
        for i, tr in enumerate(rows[:4]):
            cells = [clean(c.get_text(" ", strip=True)).lower() for c in tr.find_all(["th", "td"])]
            joined = " ".join(cells)
            if ("product" in joined or "ingredient" in joined) and (
                "volume" in joined or "amount" in joined or "inci" in joined
            ):
                header_idx, headers = i, cells
                break
        if header_idx is None:
            continue

        def col_index(terms):
            for term in terms:
                for i, h in enumerate(headers):
                    if term in h:
                        return i
            return None

        phase_i = col_index(["phase"])
        product_i = col_index(["product", "ingredient", "raw material"])
        inci_i = col_index(["inci"])
        supplier_i = col_index(["supplier"])
        amount_i = col_index(["volume", "amount", "%", "w/w"])

        if product_i is None or amount_i is None:
            continue

        current_phase = ""
        for tr in rows[header_idx + 1:]:
            cells = [clean(c.get_text(" ", strip=True)) for c in tr.find_all(["td", "th"])]
            if not cells:
                continue
            # Explicit phase-only rows update current phase.
            if phase_i is not None and phase_i < len(cells):
                ph = cells[phase_i]
                if re.fullmatch(r"[A-Z](?:\s*[-–]\s*[A-Z])?", ph):
                    current_phase = ph
            # Rows with ingredient/product cells
            product = cells[product_i] if product_i < len(cells) else ""
            if not product:
                # A few layouts place product in a bullet/list inside the row.
                product = clean(tr.get_text(" ", strip=True))
            if not product or product.lower() in {"product", "ingredient"}:
                continue
            inci = cells[inci_i] if inci_i is not None and inci_i < len(cells) else ""
            supplier = cells[supplier_i] if supplier_i is not None and supplier_i < len(cells) else ""
            amount = cells[amount_i] if amount_i < len(cells) else ""
            phase = current_phase
            if phase_i is not None and phase_i < len(cells) and cells[phase_i]:
                phase = cells[phase_i]
            # Exclude obvious non-data rows; retain source text as-is otherwise.
            if len(product) > 500:
                continue
            ingredients.append(Ingredient(
                formula_url, formula_name, phase, product, inci,
                supplier, amount, amount_to_number(amount)
            ))
        raw_tables.append(table.get_text(" ", strip=True))

    # Remove exact duplicates produced by repeated responsive tables.
    seen = set()
    unique = []
    for x in ingredients:
        key = (x.phase, x.raw_material, x.inci, x.amount_raw, x.supplier)
        if key not in seen:
            seen.add(key)
            unique.append(x)
    return unique, raw_tables

def parse_ingredient_divs(soup: BeautifulSoup, formula_url: str, formula_name: str) -> list[Ingredient]:
    """Personal Care Magazine 상세 페이지의 div 그리드 성분표.
    <div class="formulation-method"> 안에 Phase 블록(.phase)이 있고, 각 블록의
    .formulation-ingredients 아래 셀이 헤더(.method-heading) 개수만큼 반복된다."""
    out: list[Ingredient] = []
    for method in soup.select("div.formulation-method"):
        headers = [clean(h.get_text(" ", strip=True)).lower() for h in method.select(".method-heading")]
        if not headers:
            continue

        def idx(*terms):
            for t in terms:
                for i, h in enumerate(headers):
                    if t in h:
                        return i
            return None

        pi, ii, si, ai = idx("product", "ingredient"), idx("inci"), idx("supplier"), idx("volume", "amount", "%")
        if pi is None or ai is None:
            continue
        for block in method.select("div.phase"):
            grid = block.select_one(".formulation-ingredients")
            if grid is None:
                continue
            first = block.find("div", recursive=False)
            phase = clean(first.get_text(" ", strip=True)) if first and first is not grid else ""
            cells = [clean(c.get_text(" ", strip=True)) for c in grid.find_all("div", recursive=False)]
            n = len(headers)
            for k in range(0, len(cells) - n + 1, n):
                row = cells[k:k + n]
                product = row[pi]
                amount = row[ai]
                inci_v = row[ii] if ii is not None else ""
                # 상품명 없이 INCI·함량만 있는 행(예: Water / to 100)도 유지
                if not (product or inci_v or amount):
                    continue
                out.append(Ingredient(
                    formula_url, formula_name, phase, product,
                    row[ii] if ii is not None else "", row[si] if si is not None else "",
                    amount, amount_to_number(amount)))
    return out

def extract_preparation(soup: BeautifulSoup) -> str:
    # Find heading containing "Preparation" and gather following text until next major heading.
    for heading in soup.find_all(["h2", "h3", "h4", "h5", "strong", "b"]):
        if "preparation" in clean(heading.get_text(" ", strip=True)).lower():
            chunks = []
            for node in heading.find_all_next():
                if node.name in ["h1", "h2", "h3", "h4", "h5"] and node is not heading:
                    break
                if node.name in ["p", "li"]:
                    txt = clean(node.get_text(" ", strip=True))
                    if txt:
                        chunks.append(txt)
            return " ".join(chunks)[:8000]
    # fallback: page text segment after the Preparation token
    text = soup.get_text("\n", strip=True)
    m = re.search(r"Preparation\s*:?(.*?)(?:Disclaimer|$)", text, re.I | re.S)
    return clean(m.group(1))[:8000] if m else ""

def parse_detail(session: requests.Session, url: str, category: str, delay: float):
    soup = get_soup(session, url, delay)
    h1 = soup.find("h1")
    title = clean(h1.get_text(" ", strip=True)) if h1 else ""
    if not title:
        title = clean(soup.title.get_text()) if soup.title else url

    supplier = extract_supplier(soup)
    ingredients, raw_tables = parse_ingredient_table(soup, url, title)
    if not ingredients:  # 표(<table>)가 아닌 div 그리드 레이아웃
        ingredients = parse_ingredient_divs(soup, url, title)
    body_text = soup.get_text("\n", strip=True)
    ph, viscosity, appearance, stability = extract_quality(body_text)
    preparation = extract_preparation(soup)

    status = "상세 파싱 완료" if ingredients else "성분표 구조 확인 필요"
    formula = Formula(
        formula_url=url,
        formula_name=title,
        supplier_or_publisher=supplier,
        source_category=category,
        product_type=infer_product_type(title, category),
        pH=ph, viscosity=viscosity, appearance=appearance, stability=stability,
        ingredient_count=len(ingredients),
        parse_status=status,
    )
    return formula, ingredients, preparation, body_text[:15000], raw_tables

def save_excel(formulas, ingredients, processes, logs, output: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Formulation_Master"
    ws.append(list(asdict(formulas[0]).keys()) if formulas else [
        "formula_url","formula_name","supplier_or_publisher","source_category",
        "product_type","pH","viscosity","appearance","stability","ingredient_count","parse_status"
    ])
    for f in formulas:
        ws.append(list(asdict(f).values()))

    wi = wb.create_sheet("Ingredient_Detail")
    wi.append(["Formula_URL","Formula_Name","Phase","Raw_Material_as_Published",
               "INCI_as_Published","Supplier","Amount_as_Published","Amount_numeric"])
    for x in ingredients:
        wi.append([x.formula_url,x.formula_name,x.phase,x.raw_material,x.inci,
                   x.supplier,x.amount_raw,x.amount_numeric])

    wp = wb.create_sheet("Manufacturing_Process")
    wp.append(["Formula_URL","Formula_Name","Preparation_as_Published"])
    for url, name, method in processes:
        wp.append([url,name,method])

    wl = wb.create_sheet("Crawl_Log")
    wl.append(["URL","Status","Message"])
    for row in logs:
        wl.append(list(row))

    for sheet in wb.worksheets:
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for c in sheet[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="315B72")
            c.alignment = Alignment(wrap_text=True, vertical="center")
        for col in range(1, sheet.max_column + 1):
            vals = [str(sheet.cell(r,col).value or "") for r in range(1,min(sheet.max_row,150)+1)]
            sheet.column_dimensions[get_column_letter(col)].width = min(max(max(map(len,vals),default=12)+2,12),55)
        for row in sheet.iter_rows(min_row=2):
            for c in row:
                c.alignment = Alignment(vertical="top", wrap_text=True)
    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)

def save_csvs(formulas, all_ingredients, processes, logs) -> None:
    """CSV 저장. 수집 도중에도 주기적으로 호출해 중단돼도 그때까지의 결과가 남게 한다."""
    if formulas:
        with open("formulations_master.csv", "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=list(asdict(formulas[0]).keys()))
            w.writeheader()
            w.writerows(asdict(x) for x in formulas)
    if all_ingredients:
        with open("ingredients_detail.csv", "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=list(asdict(all_ingredients[0]).keys()))
            w.writeheader()
            w.writerows(asdict(x) for x in all_ingredients)
    if processes:
        with open("manufacturing_process.csv", "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["formula_url", "formula_name", "preparation"])
            w.writerows(processes)
    with open("crawl_log.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["url", "status", "message"])
        w.writerows(logs)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--categories", nargs="*", default=DEFAULT_CATEGORIES,
                        help="Category/archive page URLs")
    parser.add_argument("--urls-file", default="detail_urls.txt",
                        help="Optional text file with one detail URL per line")
    parser.add_argument("--max-pages", type=int, default=30,
                        help="수집할 처방 상세 페이지 최대 수")
    parser.add_argument("--list-pages", type=int, default=1,
                        help="카테고리별로 넘겨 볼 목록 페이지 수(?pageNumber=N). 404/새 링크 없음이면 조기 종료")
    parser.add_argument("--delay", type=float, default=3.0)
    parser.add_argument("--output", default="skincare_formulations_crawled.xlsx")
    args = parser.parse_args()

    # Explicit opt-in prevents accidental bulk collection before checking rights.
    import os
    if os.getenv("ALLOW_CRAWL") != "YES":
        raise SystemExit(
            "실행 전 사이트 이용약관/robots.txt와 대량 수집·내부 DB/AI 활용 권한을 확인하세요. "
            "확인 후 Codespaces Secrets 또는 환경변수 ALLOW_CRAWL=YES를 설정하세요."
        )

    session = make_session()
    url_category: dict[str, str] = {}   # 상세 URL → 처음 발견된 카테고리 (삽입 순서 유지)
    per_category: dict[str, list[str]] = {}
    logs = []
    blocked = False

    # Optional user-curated list of exact detail URLs
    url_file = Path(args.urls_file)
    if url_file.exists():
        manual = per_category.setdefault("Manual URL list", [])
        for line in url_file.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and line not in url_category:
                url_category[line] = "Manual URL list"
                manual.append(line)

    # Collect detail links from category/archive pages (with pagination).
    for category_url in args.categories:
        cname = category_name(category_url) or category_url
        bucket = per_category.setdefault(cname, [])
        for page in range(1, args.list_pages + 1):
            page_url = category_url if page == 1 else f"{category_url.split('?')[0]}?pageNumber={page}"
            try:
                soup = get_soup(session, page_url, args.delay)
            except PermissionError as exc:  # 차단/접근 제한 → 우회하지 않고 전체 중단
                logs.append((page_url, "BLOCKED", str(exc)))
                logging.error("접근 제한 감지, 수집 중단: %s", exc)
                blocked = True
                break
            except requests.HTTPError as exc:
                if exc.response is not None and exc.response.status_code == 404:
                    logs.append((page_url, "END", "no more list pages"))
                else:
                    logs.append((page_url, "ERROR", str(exc)))
                break
            except Exception as exc:
                logs.append((page_url, "ERROR", str(exc)))
                logging.exception("Category failed: %s", page_url)
                break
            found = collect_detail_links(soup, page_url)
            new = [u for u in found if u not in url_category]
            for u in new:
                url_category[u] = cname
                bucket.append(u)
            logs.append((page_url, "OK", f"{len(found)} detail links found, {len(new)} new"))
            if not new:
                break
        if blocked:
            break

    # 카테고리를 번갈아 가며 선택 (한 카테고리 쏠림 방지)
    detail_urls: list[str] = []
    queues = [list(v) for v in per_category.values()]
    while len(detail_urls) < args.max_pages and any(queues):
        for q in queues:
            if q and len(detail_urls) < args.max_pages:
                detail_urls.append(q.pop(0))
    logging.info("상세 페이지 %d개 수집 예정 (발견 %d개)", len(detail_urls), len(url_category))

    formulas, all_ingredients, processes = [], [], []
    for n, url in enumerate([] if blocked else detail_urls, 1):
        try:
            formula, ingredients, preparation, raw_text, raw_tables = parse_detail(
                session, url, url_category.get(url, ""), args.delay
            )
            formulas.append(formula)
            all_ingredients.extend(ingredients)
            processes.append((url, formula.formula_name, preparation))
            logs.append((url, formula.parse_status,
                         f"ingredients={len(ingredients)}; preparation={'yes' if preparation else 'no'}"))
        except PermissionError as exc:
            logs.append((url, "BLOCKED", str(exc)))
            logging.error("접근 제한 감지, 수집 중단 (%d/%d): %s", n, len(detail_urls), exc)
            break
        except Exception as exc:
            logs.append((url, "ERROR", str(exc)))
            logging.exception("Detail failed: %s", url)
        if n % 25 == 0:
            logging.info("진행 %d/%d", n, len(detail_urls))
            save_csvs(formulas, all_ingredients, processes, logs)  # 중간 저장 (중단 시 데이터 보존)

    save_excel(formulas, all_ingredients, processes, logs, Path(args.output))
    save_csvs(formulas, all_ingredients, processes, logs)

    print(f"처방 페이지: {len(formulas)}")
    print(f"성분 행: {len(all_ingredients)}")
    print(f"성분표 파싱 성공: {sum(1 for x in formulas if x.parse_status == '상세 파싱 완료')}")
    print(f"Excel 저장: {args.output}")

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    main()
