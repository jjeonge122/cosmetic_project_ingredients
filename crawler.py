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
USER_AGENT = "CosmeticFormulationResearchBot/1.0 (contact: YOUR_EMAIL)"
DEFAULT_CATEGORIES = [
    f"{BASE}/formulations-by-category/33/skin-treatment",
    f"{BASE}/formulations-by-category/15/facial-care-cleansing",
    f"{BASE}/formulations-by-category/30/anti-ageing",
    f"{BASE}/formulations-by-category/31/skin-brightening",
    f"{BASE}/formulations-by-category/22/cleansing-moisturising",
]

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
    if "captcha" in response.text.lower() or "verify you are human" in response.text.lower():
        raise PermissionError(f"자동화 차단/사람 확인 페이지 감지. 수집 중단: {url}")
    return BeautifulSoup(response.text, "html.parser")

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
            rf"{re.escape(label)}\s*[:\-]?\s*([^\n\r]+)",
            text, flags=re.IGNORECASE
        )
        if m:
            return clean(m.group(1))[:300]
    return ""

def extract_quality(text: str) -> tuple[str, str, str, str]:
    # Conservative: capture only explicit label/value pairs.
    ph = find_value_near_label(text, ["pH", "PH"])
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

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--categories", nargs="*", default=DEFAULT_CATEGORIES,
                        help="Category/archive page URLs")
    parser.add_argument("--urls-file", default="detail_urls.txt",
                        help="Optional text file with one detail URL per line")
    parser.add_argument("--max-pages", type=int, default=30)
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
    detail_urls = set()
    logs = []

    # Optional user-curated list of exact detail URLs
    url_file = Path(args.urls_file)
    if url_file.exists():
        for line in url_file.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#"):
                detail_urls.add(line)

    # Collect detail links from selected category/archive pages.
    for category_url in args.categories:
        try:
            soup = get_soup(session, category_url, args.delay)
            found = collect_detail_links(soup, category_url)
            detail_urls.update(found)
            logs.append((category_url, "OK", f"{len(found)} detail links found"))
        except Exception as exc:
            logs.append((category_url, "ERROR", str(exc)))
            logging.exception("Category failed: %s", category_url)

    detail_urls = sorted(detail_urls)[:args.max_pages]
    formulas, all_ingredients, processes = [], [], []

    for url in detail_urls:
        try:
            # Category provenance can be blank when URL came from manual list.
            formula, ingredients, preparation, raw_text, raw_tables = parse_detail(
                session, url, "", args.delay
            )
            formulas.append(formula)
            all_ingredients.extend(ingredients)
            processes.append((url, formula.formula_name, preparation))
            logs.append((url, formula.parse_status,
                         f"ingredients={len(ingredients)}; preparation={'yes' if preparation else 'no'}"))
        except Exception as exc:
            logs.append((url, "ERROR", str(exc)))
            logging.exception("Detail failed: %s", url)

    save_excel(formulas, all_ingredients, processes, logs, Path(args.output))
    # Also save machine-readable CSVs for later DB ingestion.
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

    print(f"처방 페이지: {len(formulas)}")
    print(f"성분 행: {len(all_ingredients)}")
    print(f"성분표 파싱 성공: {sum(1 for x in formulas if x.parse_status == '상세 파싱 완료')}")
    print(f"Excel 저장: {args.output}")

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    main()
