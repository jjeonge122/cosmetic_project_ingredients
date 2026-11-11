"""cas_merged.csv 의 CAS를 기준으로 화장품 원료 DB를 만든다.

가져오는 속성과 출처 (왼쪽이 우선)
  Chemical Class            화장품성분_DB.xlsx '핵심 원료'(작성자 분류) → RDKit 작용기 규칙 자동 분류
  Chemical Structure        SMILES로 RDKit 작도 (PNG)
  Molecular Formula / MW    '핵심 원료' → PubChem → RDKit 계산(SMILES만 있을 때)
  SMILES                    '핵심 원료' → PubChem
  Cosmetic Function         '핵심 원료' → 'CosIng 전체' (같은 CAS의 모든 행 합집합)
  Solubility (요약)          '핵심 원료'
  Solubility (PubChem 원문)  PubChem PUG View 'Solubility' (물 관련 문장 우선, 최대 3개)
  + MFDS 국문명·규제 요약 (cas_merged.csv)

사전 단계
  python3 merge_by_cas.py     # cas_merged.csv
  python3 pubchem.py          # data/cache/pubchem.json (약 30~40분, 이어받기 가능)

출력 (data/merged/)
  화장품원료_DB.csv            UTF-8 BOM. Chemical Structure 열 = 구조 PNG 상대경로
  화장품원료_DB.xlsx           구조 이미지 포함 (--no-xlsx 로 생략)
  structures/<CAS>.png        구조 이미지

사용법
  python3 build_ingredient_db.py [--no-xlsx]
"""

from __future__ import annotations

import argparse
import csv
import json
from collections import defaultdict
from pathlib import Path

import openpyxl
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from rdkit import Chem, RDLogger
from rdkit.Chem import Descriptors, Draw, rdMolDescriptors

from merge_ingredients import OUT_DIR, _uniq, parse_cas
from pubchem import CACHE as PUBCHEM_CACHE

RDLogger.DisableLog("rdApp.*")
XLSX_SRC = OUT_DIR / "화장품성분_DB.xlsx"
STRUCT_DIR = OUT_DIR / "structures"
J = "; "

# --------------------------------------------------------------------------- 화학 분류 (RDKit 규칙)
METALS = {"Li", "Na", "K", "Rb", "Cs", "Be", "Mg", "Ca", "Sr", "Ba", "Al", "Zn", "Fe", "Cu", "Mn", "Co",
          "Ni", "Cr", "Ti", "Zr", "Ag", "Au", "Pt", "Pb", "Sn", "Bi", "Hg", "Cd", "Se", "Sb", "As", "B"}
SMARTS = {name: Chem.MolFromSmarts(s) for name, s in {
    "carboxylic_acid": "[CX3](=O)[OX2H1]",
    "carboxylate": "[CX3](=O)[O-]",
    "aha": "[CX4]([OX2H1])[CX3](=O)[OX2H1,O-]",
    "ester": "[#6][CX3](=O)[OX2][#6]",
    "amide": "[CX3](=O)[NX3]",
    "amine": "[NX3;!$(N-[CX3]=[O,S,N]);!$(N-S(=O)=O);!$(N=*);!$(N#*);!$(N-a);!$([N]-[N])]",
    "aromatic_amine": "[NX3;H2,H1;!$(NC=O)]-a",
    "quaternary_ammonium": "[NX4+]",
    "alcohol": "[CX4][OX2H1]",
    "phenol": "c[OX2H1]",
    "aldehyde": "[CX3H1](=O)[#6]",
    "ketone": "[#6][CX3](=O)[#6]",
    "ether": "[OD2;!$(O-C=O)]([#6])[#6]",
    "sulfate_sulfonate": "[SX4](=O)(=O)[OX2H1,O-]",
    "phosphate": "[PX4](=O)([OX2])[OX2,O-]",
    "organohalide": "[#6][F,Cl,Br,I]",
    "nitro": "[N+](=O)[O-]",
    "azo": "[#6]N=N[#6]",
    "silicon": "[Si]",
    "hydroxide_ion": "[OH-]",
    "halide_ion": "[F-,Cl-,Br-,I-]",
    "sulfate_inorg": "[SX4](=O)(=O)([O-,OH])[O-,OH]",
    "carbonate": "[CX3](=O)([O-,OX2H1])[O-,OX2H1]",
    "phosphate_inorg": "[PX4](=O)([O-,OH])([O-,OH])[O-,OH]",
    "alkene": "[CX3]=[CX3]",
    "aromatic": "a",
}.items()}
GROUP_KO = {
    "carboxylic_acid": "카복실산", "carboxylate": "카복실레이트", "ester": "에스터", "amide": "아마이드",
    "amine": "아민", "aromatic_amine": "방향족 아민", "quaternary_ammonium": "4급 암모늄",
    "alcohol": "하이드록시(알코올)", "phenol": "페놀", "aldehyde": "알데하이드", "ketone": "케톤",
    "ether": "에테르", "sulfate_sulfonate": "설페이트/설포네이트", "phosphate": "인산염/에스터",
    "organohalide": "할로젠", "nitro": "나이트로", "azo": "아조", "silicon": "규소", "alkene": "알켄",
    "aromatic": "방향족 고리",
}


def classify(smiles: str) -> tuple[str, str]:
    """(주 분류, 검출 작용기). 구조 기반 근사 분류이며 작성자 수동 분류보다 우선하지 않는다."""
    m = Chem.MolFromSmiles(smiles) if smiles else None
    if m is None:
        return "", ""
    atoms = {a.GetSymbol() for a in m.GetAtoms()}
    charged = any(a.GetFormalCharge() for a in m.GetAtoms())
    g = {k for k, p in SMARTS.items() if m.HasSubstructMatch(p)}
    n_oh = len(m.GetSubstructMatches(SMARTS["alcohol"]))
    groups = ", ".join(GROUP_KO[k] for k in GROUP_KO if k in g)
    metal = atoms & METALS

    n_c = sum(a.GetSymbol() == "C" for a in m.GetAtoms())
    if "C" not in atoms or (n_c == 1 and "carbonate" in g):
        cls = ("무기 염기 (수산화물)" if "hydroxide_ion" in g and metal else
               "무기염 (할로젠화물)" if "halide_ion" in g and metal else
               "무기염 (황산염)" if "sulfate_inorg" in g and metal else
               "무기염 (인산염)" if "phosphate_inorg" in g and metal else
               "무기염 (탄산염)" if "carbonate" in g else
               "금속 산화물" if metal and atoms <= METALS | {"O"} and "O" in atoms else
               "무기 화합물 (금속 함유)" if metal else "무기 화합물")
    elif "quaternary_ammonium" in g:
        cls = "4급 암모늄염"
    elif metal and charged:
        cls = ("카복실산 금속염" if "carboxylate" in g else
               "설페이트/설포네이트 금속염" if "sulfate_sulfonate" in g else "유기 금속염")
    elif "sulfate_sulfonate" in g:
        cls = "설페이트/설포네이트"
    elif "carboxylate" in g:
        cls = "카복실산 염"
    elif "carboxylic_acid" in g and "amine" in g:
        cls = "아미노산"
    elif "carboxylic_acid" in g:
        cls = ("α-하이드록시산 (AHA)" if "aha" in g else
               "방향족 카복실산" if "aromatic" in g else "카복실산")
    elif "azo" in g:
        cls = "아조 화합물"
    elif "nitro" in g:
        cls = "나이트로 화합물"
    elif "ester" in g:
        cls = "에스터"
    elif "amide" in g:
        cls = "아마이드"
    elif "amine" in g and n_oh:
        cls = "아미노알코올"
    elif "amine" in g or "aromatic_amine" in g:
        cls = "방향족 아민" if "aromatic_amine" in g else "아민"
    elif "phenol" in g:
        cls = "페놀"
    elif "ether" in g and n_oh:
        cls = "글리콜 에테르"
    elif n_oh >= 2:
        cls = "폴리올"
    elif n_oh == 1:
        cls = "알코올"
    elif "aldehyde" in g:
        cls = "알데하이드"
    elif "ketone" in g:
        cls = "케톤"
    elif "ether" in g:
        cls = "에테르"
    elif "organohalide" in g:
        cls = "할로젠화 유기물"
    elif "silicon" in g:
        cls = "유기규소 화합물"
    elif atoms <= {"C", "H"}:
        cls = "방향족 탄화수소" if "aromatic" in g else ("불포화 탄화수소" if "alkene" in g else "포화 탄화수소")
    else:
        cls = "기타 유기 화합물"
    return cls, groups


def draw(smiles: str, path: Path) -> bool:
    m = Chem.MolFromSmiles(smiles) if smiles else None
    if m is None or m.GetNumAtoms() > 150:  # 너무 큰 분자는 생략
        return False
    if not path.exists():
        Draw.MolToFile(m, str(path), size=(240, 180))
    return True


def rdkit_formula_mw(smiles: str) -> tuple[str, str]:
    m = Chem.MolFromSmiles(smiles) if smiles else None
    if m is None:
        return "", ""
    return rdMolDescriptors.CalcMolFormula(m), f"{Descriptors.MolWt(m):.2f}"


# --------------------------------------------------------------------------- 입력
def _s(v) -> str:
    return "" if v is None or str(v).strip() in ("", "—", "-") else str(v).strip()


def load_xlsx() -> tuple[dict[str, dict], dict[str, list[dict]]]:
    wb = openpyxl.load_workbook(XLSX_SRC, read_only=True)

    def rows(sheet):
        it = wb[sheet].iter_rows(values_only=True)
        head = [str(h).strip() if h else "" for h in next(it)]
        return [dict(zip(head, r)) for r in it]

    core: dict[str, dict] = {}
    for r in rows("핵심 원료"):
        for c in parse_cas(r.get("CAS No.")):
            core.setdefault(c, r)
    cosing: dict[str, list[dict]] = defaultdict(list)
    for r in rows("CosIng 전체"):
        for c in parse_cas(r.get("CAS No.")):
            cosing[c].append(r)
    return core, cosing


def functions(rs: list[dict]) -> str:
    out = []
    for r in rs:
        for f in _s(r.get("Cosmetic Function (CosIng)")).split(","):
            if f.strip():
                out.append(f.strip())
    return ", ".join(_uniq(out))


# --------------------------------------------------------------------------- 빌드
COLUMNS = [
    "CAS No.", "국문명", "INCI Name", "Chemical Class", "Chemical Class 출처", "Chemical Structure",
    "Molecular Formula", "SMILES", "Molecular Weight (g/mol)", "구조정보 출처",
    "Cosmetic Function (CosIng)", "Solubility (요약)", "Solubility (PubChem 원문)",
    "PubChem CID", "IUPAC Name", "MFDS 원료 표준명", "MFDS 원료 영문명", "MFDS 기원및정의",
    "규제 구분", "제한 국가", "한국 제한사항", "배합금지 국가", "배합한도 국가", "MFDS 출처", "CAS 검증",
]


def build() -> tuple[list[dict], dict]:
    with open(OUT_DIR / "cas_merged.csv", encoding="utf-8-sig") as f:
        base = list(csv.DictReader(f))
    core, cosing = load_xlsx()
    pc = json.loads(PUBCHEM_CACHE.read_text(encoding="utf-8")) if PUBCHEM_CACHE.exists() else {}
    STRUCT_DIR.mkdir(parents=True, exist_ok=True)

    st = defaultdict(int)
    out = []
    for b in base:
        cas = b["CAS No."]
        k = core.get(cas, {})
        cs = cosing.get(cas, [])
        p = pc.get(cas, {})
        st["pubchem_조회됨"] += cas in pc
        st["pubchem_CID"] += bool(p.get("cid"))

        smiles = _s(k.get("SMILES")) or _s(p.get("smiles"))
        formula = _s(k.get("Molecular Formula")) or _s(p.get("formula"))
        mw = _s(k.get("Molecular Weight (g/mol)")) or _s(p.get("mw"))
        src = "핵심 원료(작성자)" if _s(k.get("SMILES")) else ("PubChem" if _s(p.get("smiles")) else "")
        if smiles and not (formula and mw):
            f2, m2 = rdkit_formula_mw(smiles)
            formula, mw = formula or f2, mw or m2
            src += " + RDKit 계산"

        cls, groups = _s(k.get("Chemical Class")), ""
        cls_src = "핵심 원료(작성자)" if cls else ""
        auto_cls, groups = classify(smiles)
        if not cls and auto_cls:
            cls, cls_src = auto_cls, f"RDKit 자동 분류 (작용기: {groups or '-'})"

        img = f"structures/{cas}.png" if smiles and draw(smiles, STRUCT_DIR / f"{cas}.png") else ""
        func = _s(k.get("Cosmetic Function (CosIng)")) or functions(cs)
        sol_pc = " | ".join((p.get("solubility") or [])[:3])
        inci = _s(k.get("INCI Name")) or J.join(_uniq(_s(r.get("INCI Name")) for r in cs))

        row = {
            "CAS No.": cas,
            "국문명": b["원료_표준명"] or b["제한_표준명"],
            "INCI Name": inci or b["원료_영문명"] or b["제한_영문명"],
            "Chemical Class": cls,
            "Chemical Class 출처": cls_src,
            "Chemical Structure": img,
            "Molecular Formula": formula,
            "SMILES": smiles,
            "Molecular Weight (g/mol)": mw,
            "구조정보 출처": src,
            "Cosmetic Function (CosIng)": func,
            "Solubility (요약)": _s(k.get("Solubility (요약)")),
            "Solubility (PubChem 원문)": sol_pc,
            "PubChem CID": p.get("cid") or "",
            "IUPAC Name": _s(p.get("iupac")),
            "MFDS 원료 표준명": b["원료_표준명"],
            "MFDS 원료 영문명": b["원료_영문명"],
            "MFDS 기원및정의": b["원료_기원및정의"],
            "규제 구분": b["제한_구분"],
            "제한 국가": b["제한_국가"],
            "한국 제한사항": b["제한_한국_제한사항"],
            "배합금지 국가": b["규제_배합금지국가"],
            "배합한도 국가": b["규제_배합한도국가"],
            "MFDS 출처": b["출처"],
            "CAS 검증": b["CAS 검증(체크섬)"],
        }
        for c in COLUMNS[3:15]:
            if row[c] and c not in ("Chemical Class 출처", "구조정보 출처"):
                st[c] += 1
        st["핵심 원료 매칭"] += bool(k)
        st["CosIng 매칭"] += bool(cs)
        out.append(row)
    st = {"CAS 수": len(out), **st}
    return out, st


def write_xlsx(rows: list[dict], path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "화장품 원료 DB"
    ws.append(COLUMNS)
    head_fill = PatternFill("solid", fgColor="FFDDEBF7")
    for c in ws[1]:
        c.font, c.fill = Font(bold=True), head_fill
        c.alignment = Alignment(wrap_text=True, vertical="center")
    widths = {"CAS No.": 13, "국문명": 22, "INCI Name": 28, "Chemical Class": 20, "Chemical Class 출처": 22,
              "Chemical Structure": 26, "SMILES": 30, "Cosmetic Function (CosIng)": 30,
              "Solubility (PubChem 원문)": 40, "IUPAC Name": 30, "한국 제한사항": 40}
    for j, name in enumerate(COLUMNS, 1):
        ws.column_dimensions[get_column_letter(j)].width = widths.get(name, 16)
    sc = COLUMNS.index("Chemical Structure") + 1
    for i, r in enumerate(rows, 2):
        for j, name in enumerate(COLUMNS, 1):
            v = r[name]
            if name == "Chemical Structure":
                continue
            if isinstance(v, str):
                v = ILLEGAL_CHARACTERS_RE.sub("", v)[:32000]
            ws.cell(i, j, v).alignment = Alignment(wrap_text=True, vertical="top")
        if r["Chemical Structure"]:
            im = XLImage(str(OUT_DIR / r["Chemical Structure"]))
            im.width, im.height = 160, 120
            ws.add_image(im, f"{get_column_letter(sc)}{i}")
            ws.row_dimensions[i].height = 92
    ws.freeze_panes = "D2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}{len(rows) + 1}"

    lg = wb.create_sheet("범례·출처")
    lg.append(["항목", "출처 / 설명"])
    for k, v in [
        ("기준", "cas_merged.csv (식약처 원료성분정보·사용제한 원료정보·규제정보를 CAS로 통합)의 CAS 1개 = 1행"),
        ("Chemical Class", "화장품성분_DB.xlsx '핵심 원료'의 작성자 분류 우선. 없으면 RDKit 작용기 규칙으로 자동 분류(근사값, 'Chemical Class 출처'에 검출 작용기 표기)"),
        ("Chemical Structure", "SMILES를 RDKit로 작도. CSV에는 structures/<CAS>.png 경로"),
        ("Molecular Formula / SMILES / MW", "'핵심 원료' 우선 → PubChem(CAS 이름 검색 첫 CID) → SMILES만 있으면 RDKit 계산. '구조정보 출처' 참고"),
        ("Cosmetic Function (CosIng)", "'핵심 원료' 우선 → 'CosIng 전체'에서 같은 CAS 행들의 기능 합집합 (EU CosIng 사본, 2023년 갱신분)"),
        ("Solubility (요약)", "'핵심 원료'의 작성자 요약 (일반 화학 지식 기반 대략값)"),
        ("Solubility (PubChem 원문)", "PubChem PUG View 'Solubility' 항목 원문, 물 관련 문장 우선 최대 3개"),
        ("PubChem 주의", "CAS로 이름 검색한 첫 CID 사용. 혼합물·고분자·식물추출물은 PubChem에 단일 화합물이 없어 비어 있는 경우가 많음"),
        ("규제 열", "식약처 사용제한 원료정보·규제정보 (cas_merged.csv)"),
    ]:
        lg.append([k, v])
    lg.column_dimensions["A"].width, lg.column_dimensions["B"].width = 28, 110
    wb.save(path)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--no-xlsx", action="store_true")
    a = ap.parse_args()
    rows, st = build()
    with open(OUT_DIR / "화장품원료_DB.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=COLUMNS)
        w.writeheader()
        w.writerows(rows)
    print(f"출력 → {OUT_DIR / '화장품원료_DB.csv'}")
    if not a.no_xlsx:
        write_xlsx(rows, OUT_DIR / "화장품원료_DB.xlsx")
        print(f"출력 → {OUT_DIR / '화장품원료_DB.xlsx'}")
    (OUT_DIR / "화장품원료_DB_report.json").write_text(json.dumps(st, ensure_ascii=False, indent=1), encoding="utf-8")
    print(json.dumps(st, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
