"""식약처 화장품 OpenAPI 3종을 'CAS No.' 하나를 기준으로 통합한다 (CAS 1개 = 1행).

입력 (merge_ingredients.py fetch 결과)
  data/raw/ingredient.json   화장품 원료성분정보   — CAS_NO 있음
  data/raw/restriction.json  화장품 사용제한 원료정보 — CAS_NO 있음
  data/raw/regulation.json   화장품 규제정보       — CAS_NO 없음
     → 사용제한 원료정보와 (표준명, 영문명)이 같은 레코드의 CAS를 물려받아 연결

규칙
  - 한 레코드에 CAS가 여러 개면('79-09-4, 137-40-6') 각 CAS 행에 모두 들어간다.
    이런 레코드는 *_multi_cas 열에 표시된다 (예: 'X 및 그 염류').
  - CAS가 같으면 같은 물질로 본다. 이름(INCI/표준명)은 비교하지 않는다.
  - CAS가 없는 레코드는 CAS 통합 표에서 빠지고 no_cas_records.csv 에 따로 남긴다.

출력 (data/merged/, UTF-8 BOM CSV)
  cas_merged.csv              CAS 1개 = 1행, 세 출처 요약
  cas_restrictions_long.csv   CAS × 사용제한(국가별) 1건 = 1행 상세
  no_cas_records.csv          CAS가 없어 통합되지 못한 레코드
  cas_merge_report.json       통계

사용법
  python3 merge_by_cas.py
"""

from __future__ import annotations

import csv
import json
from collections import defaultdict
from pathlib import Path

from merge_ingredients import (OUT_DIR, RAW_DIR, _uniq, cas_checksum_ok, clean, norm_inci, norm_kor,
                               parse_cas, split_multi)

J = "; "


def load(name: str) -> list[dict]:
    return json.loads((RAW_DIR / f"{name}.json").read_text(encoding="utf-8"))


def gkey(r: dict) -> tuple[str, str]:
    return norm_kor(r.get("INGR_STD_NAME")), norm_inci(r.get("INGR_ENG_NAME"))


def write_csv(path: Path, rows: list[dict]) -> None:
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        if not rows:
            return
        w = csv.DictWriter(f, fieldnames=list(rows[0]))
        w.writeheader()
        w.writerows(rows)


def main() -> None:
    ing, rst, reg = load("ingredient"), load("restriction"), load("regulation")

    by_cas: dict[str, dict[str, list]] = defaultdict(lambda: {"ing": [], "rst": [], "reg": []})
    no_cas: list[dict] = []

    # 1) 원료성분정보
    for r in ing:
        cas = parse_cas(r.get("CAS_NO"))
        if not cas:
            no_cas.append({"source": "원료성분정보", "std_name": clean(r.get("INGR_KOR_NAME")),
                           "eng_name": clean(r.get("INGR_ENG_NAME")), "cas_raw": clean(r.get("CAS_NO"))})
        for c in cas:
            by_cas[c]["ing"].append((r, len(cas) > 1))

    # 2) 사용제한 원료정보 + (표준명, 영문명) → CAS 사전 (규제정보 연결용)
    key_cas: dict[tuple[str, str], list[str]] = defaultdict(list)
    for r in rst:
        cas = parse_cas(r.get("CAS_NO"))
        k = gkey(r)
        for c in cas:
            if c not in key_cas[k]:
                key_cas[k].append(c)
    no_cas_rst_keys = set()
    for r in rst:
        cas = parse_cas(r.get("CAS_NO"))
        if not cas:
            k = gkey(r)
            if k not in no_cas_rst_keys and not key_cas.get(k):
                no_cas_rst_keys.add(k)
                no_cas.append({"source": "사용제한 원료정보", "std_name": clean(r.get("INGR_STD_NAME")),
                               "eng_name": clean(r.get("INGR_ENG_NAME")), "cas_raw": clean(r.get("CAS_NO"))})
        for c in cas:
            by_cas[c]["rst"].append((r, len(cas) > 1))

    # 3) 규제정보: CAS 필드 없음 → 사용제한과 같은 (표준명, 영문명)의 CAS를 물려받음
    reg_linked = 0
    for r in reg:
        cas = key_cas.get(gkey(r), [])
        if not cas:
            no_cas.append({"source": "규제정보", "std_name": clean(r.get("INGR_STD_NAME")),
                           "eng_name": clean(r.get("INGR_ENG_NAME")), "cas_raw": "(API에 CAS 필드 없음)"})
            continue
        reg_linked += 1
        for c in cas:
            by_cas[c]["reg"].append((r, len(cas) > 1))

    # --- 요약 표
    rows, long_rows = [], []
    for c in sorted(by_cas, key=lambda x: [int(p) for p in x.split("-")]):
        d = by_cas[c]
        I, R, G = d["ing"], d["rst"], d["reg"]
        kr = _uniq(clean(r.get("LIMIT_COND")) or clean(r.get("PROVIS_ATRCL")) for r, _ in R
                   if clean(r.get("COUNTRY_NAME")) in ("한국", "대한민국"))
        rows.append({
            "CAS No.": c,
            "CAS 검증(체크섬)": "OK" if cas_checksum_ok(c) else "불일치",
            "출처": "+".join(n for n, x in (("원료성분", I), ("사용제한", R), ("규제", G)) if x),
            # 원료성분정보
            "원료_표준명": J.join(_uniq(clean(r.get("INGR_KOR_NAME")) for r, _ in I)),
            "원료_영문명": J.join(_uniq(clean(r.get("INGR_ENG_NAME")) for r, _ in I)),
            "원료_기원및정의": " | ".join(_uniq(clean(r.get("ORIGIN_MAJOR_KOR_NAME")) for r, _ in I)),
            "원료_이명": " | ".join(_uniq(clean(r.get("INGR_SYNONYM")) for r, _ in I)),
            "원료_건수": len(I),
            "원료_multi_cas": "Y" if any(m for _, m in I) else "",
            # 사용제한 원료정보
            "제한_표준명": J.join(_uniq(clean(r.get("INGR_STD_NAME")) for r, _ in R)),
            "제한_영문명": J.join(_uniq(clean(r.get("INGR_ENG_NAME")) for r, _ in R)),
            "제한_구분": J.join(_uniq(clean(r.get("REGULATE_TYPE")) for r, _ in R)),
            "제한_국가": J.join(_uniq(clean(r.get("COUNTRY_NAME")) for r, _ in R)),
            "제한_한국_제한사항": " || ".join(kr),
            "제한_건수": len(R),
            "제한_multi_cas": "Y" if any(m for _, m in R) else "",
            # 규제정보
            "규제_배합금지국가": J.join(_uniq(x for r, _ in G for x in split_multi(r.get("PROH_NATIONAL")))),
            "규제_배합한도국가": J.join(_uniq(x for r, _ in G for x in split_multi(r.get("LIMIT_NATIONAL")))),
            "규제_건수": len(G),
        })
        for r, m in R:
            long_rows.append({
                "CAS No.": c, "원료_표준명": rows[-1]["원료_표준명"],
                "제한_표준명": clean(r.get("INGR_STD_NAME")), "제한_영문명": clean(r.get("INGR_ENG_NAME")),
                "구분": clean(r.get("REGULATE_TYPE")), "국가": clean(r.get("COUNTRY_NAME")),
                "고시원료명": clean(r.get("NOTICE_INGR_NAME")), "단서조항": clean(r.get("PROVIS_ATRCL")),
                "제한사항": clean(r.get("LIMIT_COND")), "원본_CAS": clean(r.get("CAS_NO")),
                "multi_cas": "Y" if m else "",
            })

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    write_csv(OUT_DIR / "cas_merged.csv", rows)
    write_csv(OUT_DIR / "cas_restrictions_long.csv", long_rows)
    write_csv(OUT_DIR / "no_cas_records.csv", no_cas)

    combos: dict[str, int] = defaultdict(int)
    for r in rows:
        combos[r["출처"]] += 1
    report = {
        "입력": {"원료성분정보": len(ing), "사용제한 원료정보": len(rst), "규제정보": len(reg)},
        "CAS 보유 레코드": {
            "원료성분정보": sum(bool(parse_cas(r.get("CAS_NO"))) for r in ing),
            "사용제한 원료정보": sum(bool(parse_cas(r.get("CAS_NO"))) for r in rst),
            "규제정보(사용제한에서 CAS 연결)": reg_linked,
        },
        "고유 CAS 수": len(rows),
        "출처 조합별 CAS 수": dict(sorted(combos.items(), key=lambda x: -x[1])),
        "체크섬 불일치 CAS": sum(r["CAS 검증(체크섬)"] != "OK" for r in rows),
        "CAS 없는 레코드(no_cas_records.csv)": len(no_cas),
    }
    (OUT_DIR / "cas_merge_report.json").write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=1))
    print(f"출력 → {OUT_DIR}/cas_merged.csv, cas_restrictions_long.csv, no_cas_records.csv")


if __name__ == "__main__":
    main()
