"""화장품성분_DB.xlsx(EU CosIng 기반)에 식약처 통합 데이터를 INCI / CAS No 기준으로 붙인다.

입력
  data/merged/화장품성분_DB.xlsx           (시트: 핵심 원료, CosIng 전체, 범례·출처)
  data/merged/ingredients_merged.csv       (merge_ingredients.py 결과)

출력 (기본 CSV, UTF-8 BOM)
  data/merged/화장품성분_DB_통합_핵심원료.csv / _CosIng전체.csv / _MFDS전용성분.csv / _범례출처.csv
  (--format xlsx 이면 data/merged/화장품성분_DB_통합.xlsx, 원본은 변경하지 않음)
  - '핵심 원료', 'CosIng 전체' 시트 오른쪽에 MFDS 열 추가
  - '핵심 원료'의 빈 '성분명(국문)' 칸을 식약처 원료성분정보 표준명으로 채움
  - 'MFDS 전용 성분' 시트: 엑셀 어느 행과도 매칭되지 않은 식약처 성분
  - '범례·출처' 시트에 추가 열 설명 기록

매칭 규칙 (엑셀 행 → 식약처 성분) — CAS 최우선, CAS가 다르면 다른 물질로 본다
  - CAS가 같은 식약처 성분은 모두 같은 물질로 붙인다.
  - INCI(정확 또는 괄호 제외)가 같은 성분은 CAS가 충돌하지 않을 때만 붙인다
    (식약처 원료성분정보는 대부분 CAS 미기재 → 이 경우 INCI로 인정).
    양쪽 CAS가 모두 있는데 다르면 다른 물질로 보고 제외(inci_rejected_cas_conflict).
  매칭 방식 표기:
    cas+inci   CAS·INCI 모두 일치한 성분이 있음
    cas        CAS만 일치 (INCI 표기 다름)
    inci / inci_loose   CAS로 찾지 못해 INCI로만 매칭 (CAS 충돌 없음)
    'cas, inci'처럼 쉼표로 이어지면: CAS 일치 성분 + CAS 미기재 동명 성분을 함께 붙임
  성분명(국문): INCI까지 일치한 원료성분정보 표준명 우선.
  INCI 비교: 정규화(대소문자·공백·대시) 정확 일치 → 괄호 속 일반명 제외 일치.
  식약처 영문명의 쉼표 별칭('Water,Aqua')도 각각 비교.

사용법
  python3 merge_xlsx.py                    # 시트별 CSV 출력 (기본)
  python3 merge_xlsx.py --format xlsx      # 엑셀 출력 (구조 이미지 보존)
  python3 merge_xlsx.py --format both
"""

from __future__ import annotations

import argparse
import copy
import csv
import json
import re
from pathlib import Path

import openpyxl
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.utils import get_column_letter

from merge_ingredients import OUT_DIR, _uniq, norm_inci, parse_cas

MAX_CELL = 4000  # 엑셀 셀 한도(32,767자)보다 훨씬 짧게 — 가독성용

MFDS_COLS = [
    ("MFDS 매칭", "식약처 성분과의 매칭 방식(CAS 우선): cas+inci=CAS·INCI 모두 일치 / cas=CAS만 일치 / inci·inci_loose=CAS 충돌 없이 INCI만 일치 / 'cas, inci'=CAS 일치 성분+CAS 미기재 동명 성분 / 빈칸=매칭 없음(CAS가 다르면 INCI가 같아도 제외). 괄호 안 숫자는 매칭된 식약처 성분 수"),
    ("MFDS 표준명", "식약처 원료성분정보 표준명(국문). 원료성분정보에 없는 규제 전용 성분이면 사용제한 원료정보의 표준명"),
    ("MFDS 영문명", "식약처 데이터의 영문명"),
    ("MFDS CAS No.", "식약처 데이터의 CAS No."),
    ("MFDS 기원 및 정의", "식약처 원료성분정보 '기원및정의'"),
    ("규제 구분", "식약처 화장품 사용제한 원료정보의 구분(금지 / 한도 / 한도/금지)"),
    ("배합금지 국가", "식약처 화장품 규제정보 PROH_NATIONAL"),
    ("배합한도 국가", "식약처 화장품 규제정보 LIMIT_NATIONAL"),
    ("한국 규제 내용", "사용제한 원료정보 중 국가=한국 행의 제한사항/단서조항 (길면 잘림)"),
    ("MFDS ID", "ingredients_merged.csv 의 id (M=원료성분정보, X=사용제한 원료정보 전용)"),
]


def txt(v: str) -> str:
    v = ILLEGAL_CHARACTERS_RE.sub("", v or "")
    return v if len(v) <= MAX_CELL else v[: MAX_CELL - 1] + "…"


def load_mfds_csv(path: str) -> list[dict]:
    with open(path, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    for r in rows:
        r["source_master"] = r["source_master"] == "True"
        r["cas_list"] = [c for x in r["cas_nos"].split(";") for c in parse_cas(x)]
    return rows


def _names(it: dict) -> list[str]:
    """식약처 영문명. 원료성분정보는 'Water,Aqua'처럼 쉼표(공백 없음+영문자)로 별칭을 나열한다."""
    full = it["inci_name"] or ""
    parts = [full]
    if it["source_master"]:
        parts += re.split(r",(?=[A-Za-z])", full)
    return _uniq(norm_inci(x) for x in parts)


def _loose(k: str) -> str:
    """괄호 속 일반명 제거: 'HEDERA HELIX (IVY) LEAF EXTRACT' → 'HEDERA HELIX LEAF EXTRACT'."""
    return re.sub(r"\s+", " ", re.sub(r"\([^()]*\)", " ", k)).strip()


class MfdsIndex:
    def __init__(self, items: list[dict]):
        self.items = items
        self.by_inci: dict[str, list[int]] = {}
        self.by_loose: dict[str, list[int]] = {}
        self.by_cas: dict[str, list[int]] = {}
        for n, it in enumerate(items):
            for k in _names(it):
                if not k:
                    continue
                self.by_inci.setdefault(k, []).append(n)
                if it["source_master"] and _loose(k) and _loose(k) != k:
                    self.by_loose.setdefault(_loose(k), []).append(n)
            for c in it["cas_list"]:
                self.by_cas.setdefault(c, []).append(n)

    def _name_hits(self, inci) -> tuple[list[int], str]:
        k = norm_inci(inci)
        if k and k in self.by_inci:
            return self.by_inci[k], "inci"
        lk = _loose(k)
        if lk and lk in self.by_loose:
            return self.by_loose[lk], "inci_loose"
        return [], ""

    def match(self, inci, cas_cell) -> tuple[list[int], str, set[int]]:
        """반환: (매칭된 식약처 성분, 매칭 방식, 이름까지 일치한 성분 집합)."""
        row_cas = parse_cas(cas_cell)
        cas_hits: list[int] = []
        for c in row_cas:
            for n in self.by_cas.get(c, []):
                if n not in cas_hits:
                    cas_hits.append(n)
        name_hits, name_how = self._name_hits(inci)
        # 이름이 같아도 양쪽 CAS가 모두 있는데 다르면 다른 물질 → 제외
        name_ok = [n for n in name_hits
                   if n in cas_hits or not row_cas or not self.items[n]["cas_list"]]
        conflict = len(name_hits) > len(name_ok)
        hits = cas_hits + [n for n in name_ok if n not in cas_hits]
        if not hits:
            return [], ("inci_rejected_cas_conflict" if conflict else ""), set()
        both = any(n in cas_hits for n in name_ok)
        extra = any(n not in cas_hits for n in name_ok)
        if cas_hits:
            how = "cas+inci" if both else "cas"
            if extra:
                how += ", " + name_how  # CAS 성분 + (CAS 미기재) 동명 성분
        else:
            how = name_how
        return hits, how, set(name_ok)


def summarize(items: list[dict], how: str) -> dict[str, str]:
    if not items:
        return {}
    kr = _uniq(x[len("[한국]"):].strip() for it in items for x in it["restriction_limits"].split(" || ")
               if x.startswith("[한국]") or x.startswith("[대한민국]"))
    split = lambda v: [x.strip() for x in v.split(";") if x.strip()]
    return {
        "MFDS 매칭": f"{how} ({len(items)})" if len(items) > 1 else how,
        "MFDS 표준명": "; ".join(_uniq(it["kor_name"] for it in items)),
        "MFDS 영문명": "; ".join(_uniq(it["inci_name"] for it in items)),
        "MFDS CAS No.": "; ".join(_uniq(c for it in items for c in it["cas_list"])),
        "MFDS 기원 및 정의": " | ".join(_uniq(it["origin_definition"] for it in items)),
        "규제 구분": "; ".join(_uniq(x for it in items for x in split(it["restriction_types"]))),
        "배합금지 국가": ", ".join(_uniq(x for it in items for x in split(it["prohibited_countries"]))),
        "배합한도 국가": ", ".join(_uniq(x for it in items for x in split(it["limited_countries"]))),
        "한국 규제 내용": " || ".join(kr),
        "MFDS ID": "; ".join(it["id"] for it in items),
    }


def header_map(ws) -> dict[str, int]:
    return {str(c.value).strip(): c.column for c in ws[1] if c.value is not None}


def add_columns(ws, idx: MfdsIndex, used: set[int], fill_kor: bool) -> dict[str, int]:
    hm = header_map(ws)
    c_inci, c_cas = hm["INCI Name"], hm["CAS No."]
    c_kor = hm.get("성분명(국문)") if fill_kor else None
    start = ws.max_column + 1
    tmpl = ws.cell(1, ws.max_column)
    for j, (name, _) in enumerate(MFDS_COLS):
        cell = ws.cell(1, start + j, name)
        cell.font, cell.fill = copy.copy(tmpl.font), copy.copy(tmpl.fill)
        cell.alignment, cell.border = copy.copy(tmpl.alignment), copy.copy(tmpl.border)
        ws.column_dimensions[get_column_letter(start + j)].width = 40 if j >= 4 else 22

    stats = {"rows": 0, "cas+inci": 0, "cas": 0, "inci": 0, "inci_loose": 0,
             "inci_rejected_cas_conflict": 0, "none": 0, "kor_filled": 0, "regulated": 0}
    for r in range(2, ws.max_row + 1):
        inci, cas = ws.cell(r, c_inci).value, ws.cell(r, c_cas).value
        if inci is None and cas is None:
            continue
        stats["rows"] += 1
        hits, how, name_ok = idx.match(inci, cas)
        stats[how or "none"] = stats.get(how or "none", 0) + 1
        if how == "inci_rejected_cas_conflict":
            how = ""
        used.update(hits)
        items = [idx.items[n] for n in hits]
        s = summarize(items, how)
        if s.get("규제 구분"):
            stats["regulated"] += 1
        for j, (name, _) in enumerate(MFDS_COLS):
            if s.get(name):
                ws.cell(r, start + j, txt(s[name]))
        if c_kor and not ws.cell(r, c_kor).value:
            # 국문명: INCI까지 일치한 원료성분정보 이름 우선, 없으면 CAS로 걸린 원료성분정보 이름
            pref = [idx.items[n] for n in hits if n in name_ok and idx.items[n]["source_master"]]
            kor = "; ".join(_uniq(it["kor_name"] for it in (pref or [it for it in items if it["source_master"]])))
            if kor:
                ws.cell(r, c_kor, kor)
                stats["kor_filled"] += 1
    if ws.auto_filter.ref:
        ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    return stats


def add_unmatched_sheet(wb, idx: MfdsIndex, used: set[int]) -> int:
    ws = wb.create_sheet("MFDS 전용 성분")
    head = ["MFDS ID", "출처", "MFDS 표준명", "MFDS 영문명", "MFDS CAS No.",
            "규제 구분", "배합금지 국가", "배합한도 국가", "한국 규제 내용", "MFDS 기원 및 정의"]
    ws.append(head)
    n = 0
    for i, it in enumerate(idx.items):
        if i in used:
            continue
        s = summarize([it], "")
        ws.append([it["id"], "원료성분정보" if it["source_master"] else "사용제한 원료정보(규제 전용)",
                   txt(it["kor_name"]), txt(it["inci_name"]), "; ".join(it["cas_list"]),
                   s["규제 구분"], s["배합금지 국가"], s["배합한도 국가"], txt(s["한국 규제 내용"]),
                   txt(it["origin_definition"])])
        n += 1
    for j, w in enumerate([10, 26, 30, 36, 18, 12, 30, 30, 50, 50], 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    for c in ws[1]:
        c.font = openpyxl.styles.Font(bold=True)
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(head))}{ws.max_row}"
    return n


CSV_NAMES = {"핵심 원료": "핵심원료", "CosIng 전체": "CosIng전체",
             "MFDS 전용 성분": "MFDS전용성분", "범례·출처": "범례출처"}


def export_csv(wb, base: Path) -> list[Path]:
    """시트별 CSV(UTF-8 BOM, 엑셀 호환). 구조 이미지는 CSV에 담을 수 없어 빠진다."""
    out = []
    for ws in wb.worksheets:
        p = base.with_name(f"{base.stem}_{CSV_NAMES.get(ws.title, ws.title)}.csv")
        with open(p, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            for row in ws.iter_rows(values_only=True):
                if any(v is not None for v in row):
                    w.writerow(["" if v is None else v for v in row])
        out.append(p)
    return out


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--xlsx", default=str(OUT_DIR / "화장품성분_DB.xlsx"))
    ap.add_argument("--mfds", default=str(OUT_DIR / "ingredients_merged.csv"))
    ap.add_argument("--out", default=str(OUT_DIR / "화장품성분_DB_통합.xlsx"))
    ap.add_argument("--format", choices=["csv", "xlsx", "both"], default="csv",
                    help="출력 형식 (기본 csv: 시트별 CSV 파일)")
    a = ap.parse_args()
    if Path(a.out).resolve() == Path(a.xlsx).resolve():
        raise SystemExit("--out 은 원본과 다른 경로여야 합니다.")

    idx = MfdsIndex(load_mfds_csv(a.mfds))
    wb = openpyxl.load_workbook(a.xlsx)  # Pillow가 설치돼 있어야 구조 이미지가 보존됨
    used: set[int] = set()
    report = {
        "핵심 원료": add_columns(wb["핵심 원료"], idx, used, fill_kor=True),
        "CosIng 전체": add_columns(wb["CosIng 전체"], idx, used, fill_kor=False),
    }
    report["MFDS 전용 성분"] = add_unmatched_sheet(wb, idx, used)
    report["MFDS 성분 중 엑셀과 매칭"] = len(used)

    lg = wb["범례·출처"]
    lg.append(["성분명(국문) (채움)", "식약처 원료성분정보 표준명으로 자동 기입 (CAS 우선, 다음 INCI 매칭). 대한화장품협회 성분사전과 다를 수 있으니 확인 필요"])
    for name, desc in MFDS_COLS:
        lg.append([name, desc])
    lg.append(["MFDS 전용 성분 시트", "식약처 데이터 중 이 엑셀의 어느 INCI/CAS와도 매칭되지 않은 성분 (주로 사용금지 물질)"])
    lg.append(["MFDS 출처", "공공데이터포털 식품의약품안전처_화장품 원료성분정보 / 사용제한 원료정보 / 규제정보 (merge_ingredients.py)"])

    if a.format in ("xlsx", "both"):
        wb.save(a.out)
        print(f"출력 → {a.out}")
    if a.format in ("csv", "both"):
        for p in export_csv(wb, Path(a.out)):
            print(f"출력 → {p}")
    print(json.dumps(report, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
