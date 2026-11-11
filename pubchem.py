"""CAS No. → PubChem 조회 (CID, 분자식, 분자량, SMILES, IUPAC명, 용해도 원문). 결과는 캐시에 누적 저장.

PubChem 이용 정책(초당 5건, 분당 400건 이하)을 지키기 위해 요청 간격을 둔다.
중단 후 다시 실행하면 캐시에 없는 것만 조회한다.

사용법
  python3 pubchem.py                # data/merged/cas_merged.csv 의 모든 CAS 조회
  python3 pubchem.py 69-72-7 56-81-5
"""

from __future__ import annotations

import csv
import json
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import requests

from merge_ingredients import OUT_DIR, ROOT

CACHE = ROOT / "data" / "cache" / "pubchem.json"
BASE = "https://pubchem.ncbi.nlm.nih.gov/rest"
INTERVAL = 0.4   # 초 — 요청 시작 간격(전체 스레드 공통), 초당 2.5건 (PubChem 한도 5건의 절반)
WORKERS = 2      # 동시 요청 수 (PubChem은 요청 처리 시간 합계도 제한하므로 낮게 유지)
PROPS = "MolecularFormula,MolecularWeight,SMILES,ConnectivitySMILES,IUPACName"

_local = threading.local()
_lock = threading.Lock()
_last = [0.0]


def _session() -> requests.Session:
    if not hasattr(_local, "s"):
        _local.s = requests.Session()
        _local.s.headers["User-Agent"] = "cosmetic-ingredient-db/1.0 (python-requests)"
    return _local.s


def _throttle() -> None:
    with _lock:
        wait = INTERVAL - (time.time() - _last[0])
        if wait > 0:
            time.sleep(wait)
        _last[0] = time.time()


class TransientError(RuntimeError):
    pass


def _get(url: str, params: dict | None = None, method: str = "GET", data: dict | None = None):
    """404 → None. 503(서버 혼잡)·네트워크 오류는 지수 백오프로 재시도."""
    for attempt in range(6):
        _throttle()
        try:
            r = _session().request(method, url, params=params, data=data, timeout=60)
        except requests.RequestException:
            time.sleep(2 ** attempt)
            continue
        if r.status_code == 404:
            return None
        if r.status_code in (429, 503):  # 과다 요청/차단 — 길게 쉬고 재시도
            time.sleep(30 * (attempt + 1))
            continue
        if r.status_code in (500, 502, 504):
            time.sleep(2 ** attempt + 1)
            continue
        if r.status_code >= 400:  # 400 등 영구 오류: 해당 항목만 건너뜀
            return None
        try:
            return r.json()
        except ValueError:
            return None
    raise TransientError(url)  # 재시도 소진 — 호출 측에서 캐시하지 않고 다음 실행 때 재조회


def load_cache() -> dict:
    return json.loads(CACHE.read_text(encoding="utf-8")) if CACHE.exists() else {}


def save_cache(c: dict) -> None:
    CACHE.parent.mkdir(parents=True, exist_ok=True)
    tmp = CACHE.with_suffix(".tmp")
    tmp.write_text(json.dumps(c, ensure_ascii=False), encoding="utf-8")
    tmp.replace(CACHE)


def _strings(o, out: list[str]) -> None:
    if isinstance(o, dict):
        if isinstance(o.get("String"), str):
            out.append(o["String"])
        for v in o.values():
            _strings(v, out)
    elif isinstance(o, list):
        for v in o:
            _strings(v, out)


def solubility(cid: int) -> list[str] | None:
    try:
        d = _get(f"{BASE}/pug_view/data/compound/{cid}/JSON", {"heading": "Solubility"})
    except TransientError:
        return None  # 캐시하지 않음
    if not d:
        return []
    out: list[str] = []
    _strings(d.get("Record", {}).get("Section", []), out)
    seen, res = set(), []
    for s in out:
        s = s.strip()
        if s and s.lower() not in seen:
            seen.add(s.lower())
            res.append(s)
    # 물에 대한 용해도를 앞으로
    res.sort(key=lambda s: 0 if "water" in s.lower() else 1)
    return res[:6]


def run(cas_list: list[str]) -> dict:
    cache = load_cache()
    todo = [c for c in cas_list if c not in cache]
    print(f"[pubchem] 대상 {len(cas_list)} / 캐시 {len(cas_list) - len(todo)} / 조회 {len(todo)}", file=sys.stderr)

    # 1) CAS → CID
    def cid_of(cas):
        try:
            d = _get(f"{BASE}/pug/compound/name/{cas}/cids/JSON")
        except TransientError:
            return cas, None
        cids = (d or {}).get("IdentifierList", {}).get("CID", [])
        return cas, {"cid": cids[0] if cids else None, "cid_count": len(cids)}

    with ThreadPoolExecutor(WORKERS) as ex:
        for i, (cas, v) in enumerate(ex.map(cid_of, todo), 1):
            if v is not None:
                cache[cas] = v
            if i % 200 == 0:
                save_cache(cache)
                print(f"  CID {i}/{len(todo)}", file=sys.stderr, flush=True)
    save_cache(cache)

    # 2) 속성 (CID 100개씩 일괄)
    need = sorted({v["cid"] for c, v in cache.items() if c in cas_list and v["cid"] and "formula" not in v})
    props: dict[int, dict] = {}
    for i in range(0, len(need), 100):
        chunk = need[i:i + 100]
        try:
            d = _get(f"{BASE}/pug/compound/cid/property/{PROPS}/JSON", method="POST",
                     data={"cid": ",".join(map(str, chunk))})
        except TransientError:
            d = None
        for p in (d or {}).get("PropertyTable", {}).get("Properties", []):
            props[p["CID"]] = p
    for c, v in cache.items():
        p = props.get(v.get("cid"))
        if p:
            v.update(formula=p.get("MolecularFormula"), mw=p.get("MolecularWeight"),
                     smiles=p.get("SMILES") or p.get("ConnectivitySMILES"), iupac=p.get("IUPACName"))
    save_cache(cache)
    print(f"[pubchem] 속성 {len(props)}건", file=sys.stderr)

    # 3) 용해도 원문 (CID별)
    need = [c for c in cas_list if c in cache and cache[c].get("cid") and "solubility" not in cache[c]]
    with ThreadPoolExecutor(WORKERS) as ex:
        for i, (cas, sol) in enumerate(ex.map(lambda c: (c, solubility(cache[c]["cid"])), need), 1):
            if sol is not None:
                cache[cas]["solubility"] = sol
            if i % 200 == 0:
                save_cache(cache)
                print(f"  용해도 {i}/{len(need)}", file=sys.stderr, flush=True)
    save_cache(cache)

    found = sum(bool(cache.get(c, {}).get("cid")) for c in cas_list)
    sol = sum(bool(cache.get(c, {}).get("solubility")) for c in cas_list)
    missing = sum(c not in cache for c in cas_list)
    if missing:
        print(f"[pubchem] 일시 오류로 미조회 {missing}건 — 다시 실행하면 이어서 조회", file=sys.stderr)
    print(f"[pubchem] 완료: CID {found}/{len(cas_list)}, 용해도 {sol}", file=sys.stderr)
    return cache


if __name__ == "__main__":
    if len(sys.argv) > 1:
        cas = sys.argv[1:]
    else:
        with open(OUT_DIR / "cas_merged.csv", encoding="utf-8-sig") as f:
            cas = [r["CAS No."] for r in csv.DictReader(f)]
    c = run(cas)
    if len(sys.argv) > 1:
        print(json.dumps({k: c[k] for k in cas}, ensure_ascii=False, indent=1))
