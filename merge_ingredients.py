"""식약처 화장품 OpenAPI 3종을 수집해 CAS No / INCI(영문명) 기준으로 통합한다.

대상 API (공공데이터포털, 제공기관 식품의약품안전처)
  1) 화장품 원료성분정보     CsmtcsIngdCpntInfoService01/getCsmtcsIngdCpntInfoService01
     INGR_KOR_NAME, INGR_ENG_NAME, CAS_NO, ORIGIN_MAJOR_KOR_NAME, INGR_SYNONYM
  2) 화장품 사용제한 원료정보  CsmtcsUseRstrcInfoService/getCsmtcsUseRstrcInfoService
     REGULATE_TYPE, INGR_STD_NAME, INGR_ENG_NAME, CAS_NO, INGR_SYNONYM,
     COUNTRY_NAME, NOTICE_INGR_NAME, PROVIS_ATRCL, LIMIT_COND
  3) 화장품 규제정보         CsmtcsReglMaterialInfoService/getCsmtcsReglMaterialInfoService
     SN, INGR_STD_NAME, INGR_ENG_NAME, PROH_NATIONAL, LIMIT_NATIONAL   (※ CAS_NO 없음)

통합 규칙
  - 원료성분정보를 마스터로 삼는다.
  - 사용제한·규제정보 레코드는 먼저 (표준명, 영문명) 정확 일치로 '규제 성분' 그룹으로 묶는다.
    (두 API는 동일한 성분 목록을 공유하며, 규제정보는 사용제한의 성분별 국가 요약이다.)
  - 각 그룹을 마스터에 매칭: INCI(영문명 정규화) → 국문 표준명 → CAS No.
    CAS가 여러 개인 그룹('X 및 그 염류')은 CAS가 걸리는 마스터 원료 모두에 붙는다(matched_count로 표시).
  - 마스터에 없는 그룹(주로 사용금지 물질)은 별도 성분(id X…, source_master=False)으로 추가한다.
    비마스터 그룹끼리는 병합하지 않는다(연쇄 병합 방지).

사용법
  export DATA_GO_KR_SERVICE_KEY='발급받은 인증키(Decoding 키 권장)'
  python3 merge_ingredients.py fetch    # data/raw/*.json + API별 원본 CSV 저장
  python3 merge_ingredients.py raw-csv  # 이미 받은 data/raw/*.json 을 CSV로만 변환
  python3 merge_ingredients.py merge    # data/merged/* 생성
  python3 merge_ingredients.py all      # fetch + merge
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import sys
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

BASE = "https://apis.data.go.kr/1471000"
SOURCES = {
    "ingredient": ("CsmtcsIngdCpntInfoService01", "getCsmtcsIngdCpntInfoService01"),
    "restriction": ("CsmtcsUseRstrcInfoService", "getCsmtcsUseRstrcInfoService"),
    "regulation": ("CsmtcsReglMaterialInfoService", "getCsmtcsReglMaterialInfoService"),
}
ROOT = Path(__file__).resolve().parent
RAW_DIR = ROOT / "data" / "raw"
OUT_DIR = ROOT / "data" / "merged"
PAGE_SIZE = 500  # API 최대값


# --------------------------------------------------------------------------- fetch
class ApiError(RuntimeError):
    pass


def _build_url(service: str, op: str, key: str, page: int, rows: int) -> str:
    # 공공데이터포털 키는 Encoding(%포함)/Decoding 두 종류가 있다. 이미 인코딩된 키는 그대로 쓴다.
    enc_key = key if "%" in key else urllib.parse.quote(key, safe="")
    qs = urllib.parse.urlencode({"pageNo": page, "numOfRows": rows, "type": "json"})
    return f"{BASE}/{service}/{op}?serviceKey={enc_key}&{qs}"


def _extract_items(payload: dict) -> tuple[list[dict], int]:
    if "OpenAPI_ServiceResponse" in payload:  # 게이트웨이 오류(인증키 등)
        hdr = payload["OpenAPI_ServiceResponse"].get("cmmMsgHeader", {})
        raise ApiError(f"{hdr.get('errMsg')} ({hdr.get('returnAuthMsg')}, code {hdr.get('returnReasonCode')})")
    resp = payload.get("response", payload)
    header = resp.get("header", {})
    code = str(header.get("resultCode", "00"))
    if code not in ("00", "0", "200"):
        raise ApiError(f"resultCode={code} {header.get('resultMsg')}")
    body = resp.get("body", {}) or {}
    items = body.get("items") or []
    if isinstance(items, dict):
        items = items.get("item") or []
    if isinstance(items, dict):
        items = [items]
    # 일부 API는 items: [{"item": {...}}] 형태
    items = [i["item"] if isinstance(i, dict) and set(i) == {"item"} else i for i in items]
    return items, int(body.get("totalCount") or 0)


def _get_json(url: str, retries: int = 4) -> dict:
    last = None
    for attempt in range(retries):
        try:
            with urllib.request.urlopen(url, timeout=60) as r:
                text = r.read().decode("utf-8")
            try:
                return json.loads(text)
            except json.JSONDecodeError:
                raise ApiError(f"JSON 아님: {text[:200]}")
        except (urllib.error.URLError, TimeoutError, ApiError) as e:
            last = e
            time.sleep(1.5 * (attempt + 1))
    raise ApiError(f"요청 실패: {last}")


def fetch_source(name: str, key: str) -> list[dict]:
    service, op = SOURCES[name]
    items, total = _extract_items(_get_json(_build_url(service, op, key, 1, PAGE_SIZE)))
    pages = (total + PAGE_SIZE - 1) // PAGE_SIZE
    print(f"[{name}] totalCount={total}, pages={pages}", file=sys.stderr)
    for page in range(2, pages + 1):
        more, _ = _extract_items(_get_json(_build_url(service, op, key, page, PAGE_SIZE)))
        items.extend(more)
        if page % 20 == 0:
            print(f"  [{name}] {page}/{pages}", file=sys.stderr)
    if total and len(items) != total:
        print(f"  경고: [{name}] 수집 {len(items)}건 ≠ totalCount {total}", file=sys.stderr)
    return items


RAW_CSV_NAMES = {
    "ingredient": "화장품_원료성분정보",
    "restriction": "화장품_사용제한_원료정보",
    "regulation": "화장품_규제정보",
}


def save_raw_csv(name: str, items: list[dict]) -> Path:
    """API 응답 원본을 그대로 CSV로 저장 (값 가공 없음, 열은 API 필드명, UTF-8 BOM)."""
    cols: list[str] = []
    for it in items:
        for k in it:
            if k not in cols:
                cols.append(k)
    path = RAW_DIR / f"{RAW_CSV_NAMES[name]}.csv"
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        for it in items:
            w.writerow({k: "" if it.get(k) is None else it[k] for k in cols})
    return path


def cmd_raw_csv() -> None:
    """이미 받아 둔 data/raw/*.json 을 CSV로 변환 (API 재호출 없음)."""
    for name in SOURCES:
        src = RAW_DIR / f"{name}.json"
        if not src.exists():
            sys.exit(f"{src} 가 없습니다. 먼저 `fetch` 를 실행하세요.")
        items = json.loads(src.read_text(encoding="utf-8"))
        print(f"[{name}] {len(items)}건 → {save_raw_csv(name, items)}", file=sys.stderr)


def cmd_fetch() -> None:
    key = os.environ.get("DATA_GO_KR_SERVICE_KEY", "").strip()
    if not key:
        sys.exit("환경변수 DATA_GO_KR_SERVICE_KEY 에 공공데이터포털 인증키를 설정하세요.")
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    for name in SOURCES:
        items = fetch_source(name, key)
        (RAW_DIR / f"{name}.json").write_text(json.dumps(items, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"[{name}] {len(items)}건 저장 → {RAW_DIR / (name + '.json')}", file=sys.stderr)
        print(f"[{name}] CSV 저장 → {save_raw_csv(name, items)}", file=sys.stderr)


# --------------------------------------------------------------------------- normalize
CAS_RE = re.compile(r"(?<!\d)(\d{2,7})\s*-\s*(\d{2})\s*-\s*(\d)(?!\d)")
DASHES = dict.fromkeys(map(ord, "‐‑‒–—―−－"), "-")


def clean(v) -> str:
    if v is None:
        return ""
    return re.sub(r"\s+", " ", unicodedata.normalize("NFKC", str(v))).strip()


def parse_cas(v) -> list[str]:
    """'8001-54-5 / 63449-41-2' 같은 다중 표기에서 CAS 번호를 모두 추출(순서 유지, 중복 제거)."""
    out = []
    for a, b, c in CAS_RE.findall(clean(v).translate(DASHES)):
        cas = f"{int(a)}-{b}-{c}"
        if cas not in out:
            out.append(cas)
    return out


def cas_checksum_ok(cas: str) -> bool:
    digits = cas.replace("-", "")
    body, check = digits[:-1], int(digits[-1])
    return sum(int(d) * (i + 1) for i, d in enumerate(reversed(body))) % 10 == check


def norm_inci(v) -> str:
    s = clean(v).translate(DASHES).upper()
    s = s.replace("’", "'").replace("‘", "'").replace("`", "'")
    s = re.sub(r"\s*([-/,()])\s*", r"\1", s)
    s = re.sub(r"[.\s]+$", "", s)
    return s


def norm_kor(v) -> str:
    s = clean(v).translate(DASHES)
    return re.sub(r"\s+", "", s)


def split_multi(v) -> list[str]:
    return [x for x in (clean(p) for p in re.split(r"[,\n;|]", clean(v))) if x]


# --------------------------------------------------------------------------- merge
class Ingredient:
    __slots__ = ("id", "kor_name", "inci_name", "cas_nos", "origin_definition", "synonyms",
                 "source_master", "restrictions", "regulations", "match_methods")

    def __init__(self, id_: str, kor: str, inci: str, cas: list[str], origin: str, syn: str, master: bool):
        self.id = id_
        self.kor_name = kor
        self.inci_name = inci
        self.cas_nos = list(cas)
        self.origin_definition = origin
        self.synonyms = syn
        self.source_master = master
        self.restrictions: list[dict] = []
        self.regulations: list[dict] = []
        self.match_methods: set[str] = set()

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "kor_name": self.kor_name,
            "inci_name": self.inci_name,
            "cas_nos": self.cas_nos,
            "cas_checksum_ok": all(cas_checksum_ok(c) for c in self.cas_nos) if self.cas_nos else None,
            "origin_definition": self.origin_definition,
            "synonyms": self.synonyms,
            "source_master": self.source_master,
            "match_methods": sorted(self.match_methods),
            "restrictions": self.restrictions,
            "regulations": self.regulations,
        }


class Index:
    def __init__(self):
        self.by_cas: dict[str, list[Ingredient]] = {}
        self.by_inci: dict[str, list[Ingredient]] = {}
        self.by_kor: dict[str, list[Ingredient]] = {}

    @staticmethod
    def _add(d, k, ing):
        if k:
            lst = d.setdefault(k, [])
            if ing not in lst:
                lst.append(ing)

    def add(self, ing: Ingredient, cas: list[str], inci: str, kor: str):
        for c in cas:
            self._add(self.by_cas, c, ing)
        self._add(self.by_inci, norm_inci(inci), ing)
        self._add(self.by_kor, norm_kor(kor), ing)

    def match(self, cas: list[str], inci: str, kor: str) -> tuple[list[Ingredient], str]:
        ni, nk = norm_inci(inci), norm_kor(kor)
        # 1) 이름(INCI → 국문 표준명) 정확 일치가 가장 확실한 동일성
        if ni and ni in self.by_inci:
            return self.by_inci[ni], "inci"
        if nk and nk in self.by_kor:
            return self.by_kor[nk], "kor_name"
        # 2) CAS 일치. 'X 및 그 염류'처럼 CAS가 여러 개인 규제 성분은 해당 염류 원료 모두에 붙는다.
        cas_hits: list[Ingredient] = []
        for c in cas:
            for ing in self.by_cas.get(c, []):
                if ing not in cas_hits:
                    cas_hits.append(ing)
        if cas_hits:
            return cas_hits, "cas"
        return [], ""


class RegGroup:
    """사용제한/규제정보 레코드를 (표준명, 영문명) 정확 일치로 묶은 '규제 성분' 단위.
    두 API는 동일한 (표준명, 영문명) 집합을 공유한다(규제정보 = 사용제한의 성분별 요약)."""

    def __init__(self, std: str, eng: str):
        self.std, self.eng = std, eng
        self.cas: list[str] = []
        self.synonyms: list[str] = []
        self.restrictions: list[dict] = []
        self.regulations: list[dict] = []


def _gkey(rec: dict) -> tuple[str, str]:
    return norm_kor(rec.get("INGR_STD_NAME")), norm_inci(rec.get("INGR_ENG_NAME"))


def merge(ingredients: list[dict], restrictions: list[dict], regulations: list[dict]) -> tuple[list[Ingredient], dict]:
    master_idx = Index()
    result: list[Ingredient] = []
    stats: dict = {"master": len(ingredients)}

    for n, r in enumerate(ingredients, 1):
        cas = parse_cas(r.get("CAS_NO"))
        ing = Ingredient(f"M{n:06d}", clean(r.get("INGR_KOR_NAME")), clean(r.get("INGR_ENG_NAME")), cas,
                         clean(r.get("ORIGIN_MAJOR_KOR_NAME")), clean(r.get("INGR_SYNONYM")), True)
        master_idx.add(ing, cas, ing.inci_name, ing.kor_name)
        result.append(ing)

    # --- 규제 성분 그룹 구성
    groups: dict[tuple[str, str], RegGroup] = {}

    def group(rec: dict) -> RegGroup:
        k = _gkey(rec)
        if k not in groups:
            groups[k] = RegGroup(clean(rec.get("INGR_STD_NAME")), clean(rec.get("INGR_ENG_NAME")))
        return groups[k]

    for r in restrictions:
        g = group(r)
        for c in parse_cas(r.get("CAS_NO")):
            if c not in g.cas:
                g.cas.append(c)
        syn = clean(r.get("INGR_SYNONYM"))
        if syn and syn not in g.synonyms:
            g.synonyms.append(syn)
        g.restrictions.append({
            "regulate_type": clean(r.get("REGULATE_TYPE")),
            "std_name": g.std, "eng_name": g.eng,
            "cas_no_raw": clean(r.get("CAS_NO")),
            "country": clean(r.get("COUNTRY_NAME")),
            "notice_ingr_name": clean(r.get("NOTICE_INGR_NAME")),
            "provision": clean(r.get("PROVIS_ATRCL")),
            "limit_cond": clean(r.get("LIMIT_COND")),
        })
    for r in regulations:  # 규제정보 API에는 CAS_NO가 없다 → 그룹 키로만 연결
        g = group(r)
        g.regulations.append({
            "sn": clean(r.get("SN")),
            "std_name": g.std, "eng_name": g.eng,
            "prohibited_countries": split_multi(r.get("PROH_NATIONAL")),
            "limited_countries": split_multi(r.get("LIMIT_NATIONAL")),
        })

    # --- 그룹 → 마스터 매칭
    by_how: dict[str, int] = {}
    multi = 0
    extra_seq = 0
    for g in groups.values():
        hits, how = master_idx.match(g.cas, g.eng, g.std)
        if not hits:
            extra_seq += 1
            how = "unmatched"
            hits = [Ingredient(f"X{extra_seq:06d}", g.std, g.eng, g.cas, "", " | ".join(g.synonyms), False)]
            result.append(hits[0])
        by_how[how] = by_how.get(how, 0) + 1
        multi += len(hits) > 1
        for h in hits:
            h.match_methods.add(how)
            extra = {"match": how, "matched_count": len(hits)}
            h.restrictions.extend({**x, **extra} for x in g.restrictions)
            h.regulations.extend({**x, **extra} for x in g.regulations)

    stats["reg_groups"] = len(groups)
    stats["reg_groups_by_match"] = by_how
    stats["reg_groups_matched_to_multiple_masters"] = multi
    stats["restriction_rows"] = len(restrictions)
    stats["regulation_rows"] = len(regulations)
    stats["groups_without_regulation_row"] = sum(not g.regulations for g in groups.values())
    stats["groups_without_restriction_row"] = sum(not g.restrictions for g in groups.values())
    return result, stats


def _uniq(xs):
    seen, out = set(), []
    for x in xs:
        if x and x not in seen:
            seen.add(x)
            out.append(x)
    return out


def flat_row(ing: Ingredient) -> dict:
    rs, gs = ing.restrictions, ing.regulations
    return {
        "id": ing.id,
        "kor_name": ing.kor_name,
        "inci_name": ing.inci_name,
        "cas_nos": "; ".join(ing.cas_nos),
        "source_master": ing.source_master,
        "origin_definition": ing.origin_definition,
        "synonyms": ing.synonyms,
        "has_restriction": bool(rs),
        "restriction_types": "; ".join(_uniq(r["regulate_type"] for r in rs)),
        "restriction_countries": "; ".join(_uniq(r["country"] for r in rs)),
        "restriction_limits": " || ".join(_uniq(
            f"[{r['country']}] {r['limit_cond']}".strip() for r in rs if r["limit_cond"])),
        "has_regulation": bool(gs),
        "prohibited_countries": "; ".join(_uniq(c for g in gs for c in g["prohibited_countries"])),
        "limited_countries": "; ".join(_uniq(c for g in gs for c in g["limited_countries"])),
        "match_methods": "; ".join(sorted(ing.match_methods)),
    }


def cmd_merge() -> None:
    def load(name):
        p = RAW_DIR / f"{name}.json"
        if not p.exists():
            sys.exit(f"{p} 가 없습니다. 먼저 `fetch` 를 실행하세요.")
        return json.loads(p.read_text(encoding="utf-8"))

    result, stats = merge(load("ingredient"), load("restriction"), load("regulation"))
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    (OUT_DIR / "ingredients_merged.json").write_text(
        json.dumps([i.to_dict() for i in result], ensure_ascii=False, indent=1), encoding="utf-8")

    rows = [flat_row(i) for i in result]
    with open(OUT_DIR / "ingredients_merged.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()) if rows else ["id"])
        w.writeheader()
        w.writerows(rows)

    # 제한/규제 1건 = 1행 (long format) — 분석·필터링용
    with open(OUT_DIR / "restrictions_long.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = None
        for i in result:
            for kind, recs in (("restriction", i.restrictions), ("regulation", i.regulations)):
                for r in recs:
                    row = {"id": i.id, "kor_name": i.kor_name, "inci_name": i.inci_name,
                           "cas_nos": "; ".join(i.cas_nos), "source": kind,
                           "regulate_type": r.get("regulate_type", ""),
                           "country": r.get("country", ""),
                           "prohibited_countries": "; ".join(r.get("prohibited_countries", [])),
                           "limited_countries": "; ".join(r.get("limited_countries", [])),
                           "notice_ingr_name": r.get("notice_ingr_name", ""),
                           "provision": r.get("provision", ""), "limit_cond": r.get("limit_cond", ""),
                           "match": r["match"], "matched_count": r["matched_count"]}
                    if w is None:
                        w = csv.DictWriter(f, fieldnames=list(row))
                        w.writeheader()
                    w.writerow(row)

    stats["output_ingredients"] = len(result)
    stats["non_master_ingredients"] = sum(not i.source_master for i in result)
    (OUT_DIR / "merge_report.json").write_text(json.dumps(stats, ensure_ascii=False, indent=1), encoding="utf-8")
    print(json.dumps(stats, ensure_ascii=False, indent=1))
    print(f"출력 → {OUT_DIR}", file=sys.stderr)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("command", choices=["fetch", "merge", "all", "raw-csv"])
    a = ap.parse_args()
    if a.command == "raw-csv":
        cmd_raw_csv()
    if a.command in ("fetch", "all"):
        cmd_fetch()
    if a.command in ("merge", "all"):
        cmd_merge()


if __name__ == "__main__":
    main()
