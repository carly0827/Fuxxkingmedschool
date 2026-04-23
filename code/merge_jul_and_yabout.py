from __future__ import annotations

import argparse
import json
import re
import unicodedata
import zipfile
from dataclasses import dataclass, asdict
from difflib import SequenceMatcher
from pathlib import Path
from typing import Dict, List, Tuple

import fitz  # PyMuPDF


@dataclass
class PageInfo:
    page_no: int  # 1-based
    raw_text: str
    norm_text: str


@dataclass
class PageRef:
    source: str  # "jul" | "yabout"
    page_no: int  # 1-based
    reason: str  # original_jul | inserted_yabout_front | inserted_yabout_middle | inserted_yabout_tail


def normalize_text(text: str) -> str:
    text = unicodedata.normalize("NFKC", text or "")
    text = text.replace("\x00", " ")
    text = text.lower()
    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(r"[^0-9a-z가-힣]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def extract_page_infos(pdf_path: Path) -> List[PageInfo]:
    doc = fitz.open(pdf_path)
    infos: List[PageInfo] = []
    try:
        for i, page in enumerate(doc):
            raw = page.get_text("text") or ""
            infos.append(PageInfo(page_no=i + 1, raw_text=raw, norm_text=normalize_text(raw)))
    finally:
        doc.close()
    return infos


def similarity(a: str, b: str) -> float:
    if not a and not b:
        return 1.0
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    ratio = SequenceMatcher(None, a[:5000], b[:5000]).ratio()
    ta = set(a.split())
    tb = set(b.split())
    jaccard = (len(ta & tb) / len(ta | tb)) if (ta and tb) else 0.0
    return max(ratio, 0.7 * ratio + 0.3 * jaccard)


def align_pages(
    jul_pages: List[PageInfo],
    yabout_pages_excluding_first: List[PageInfo],
    lookahead: int = 30,
    threshold: float = 0.92,
    fallback_threshold: float = 0.80,
) -> Tuple[List[Tuple[int, int, float]], List[int]]:
    """Return matches between jul[1:] and yabout[1:], and unmatched yabout pages."""
    matches: List[Tuple[int, int, float]] = []
    extras: List[int] = []

    j = 0
    total_y = len(yabout_pages_excluding_first)

    for i, jul in enumerate(jul_pages):
        best_idx = None
        best_score = -1.0
        for k in range(j, min(total_y, j + lookahead)):
            score = similarity(jul.norm_text, yabout_pages_excluding_first[k].norm_text)
            if score > best_score:
                best_score = score
                best_idx = k
            if score == 1.0:
                break

        if best_score < threshold:
            for k in range(min(total_y, j + lookahead), total_y):
                raw_score = similarity(jul.norm_text, yabout_pages_excluding_first[k].norm_text)
                adjusted = raw_score - 0.003 * (k - j)
                if adjusted > best_score:
                    best_score = adjusted
                    best_idx = k

        if best_idx is None or best_score < fallback_threshold:
            raise RuntimeError(
                f"줄 파일 {i+2}페이지와 대응되는 야붙 페이지를 찾지 못했습니다. "
                f"(best_score={best_score:.3f})"
            )

        for extra_idx in range(j, best_idx):
            extras.append(extra_idx)

        real_score = similarity(jul.norm_text, yabout_pages_excluding_first[best_idx].norm_text)
        matches.append((i, best_idx, real_score))
        j = best_idx + 1

    for extra_idx in range(j, total_y):
        extras.append(extra_idx)

    return matches, extras


def _nfc(s: str) -> str:
    return unicodedata.normalize('NFC', s)


def build_output_stem(jul_path: Path) -> str:
    stem = re.sub(r"\s*\(\d+\)$", "", _nfc(jul_path.stem))
    if re.search(r"줄$", stem):
        return re.sub(r"줄$", "줄과 야붙", stem)
    return f"{stem}_줄과 야붙"


def auto_detect(folder: Path) -> Tuple[Path, Path]:
    pdfs = sorted(folder.glob("*.pdf"))
    jul_candidates = [p for p in pdfs if "줄" in _nfc(p.stem) and "야붙" not in _nfc(p.stem)]
    yabout_candidates = [p for p in pdfs if "야붙" in _nfc(p.stem)]
    if len(jul_candidates) != 1 or len(yabout_candidates) != 1:
        raise RuntimeError(
            "자동 탐지에 실패했습니다. --jul 과 --yabout 경로를 직접 넣어 주세요.\n"
            f"줄 후보: {[p.name for p in jul_candidates]}\n"
            f"야붙 후보: {[p.name for p in yabout_candidates]}"
        )
    return jul_candidates[0], yabout_candidates[0]


def compute_page_plan(jul_path: Path, yabout_path: Path) -> Tuple[List[PageRef], dict]:
    jul_infos = extract_page_infos(jul_path)
    yabout_infos = extract_page_infos(yabout_path)
    if not jul_infos:
        raise RuntimeError("줄 파일이 비어 있습니다.")
    if not yabout_infos:
        raise RuntimeError("야붙 파일이 비어 있습니다.")

    matches_tail, _extra_indices_excl_first = align_pages(jul_infos[1:], yabout_infos[1:])
    matches = [(jul_idx0 + 1, y_idx0_excl_first, score) for jul_idx0, y_idx0_excl_first, score in matches_tail]

    plan: List[PageRef] = [PageRef("yabout", 1, "inserted_yabout_front"), PageRef("jul", 1, "original_jul")]
    prev_y_idx_excl_first = 0
    low_conf = []

    for jul_idx0, y_idx0_excl_first, score in matches:
        for k in range(prev_y_idx_excl_first, y_idx0_excl_first):
            abs_page_1based = k + 2
            plan.append(PageRef("yabout", abs_page_1based, "inserted_yabout_middle"))
        abs_jul_page = jul_idx0 + 1
        plan.append(PageRef("jul", abs_jul_page, "original_jul"))
        if score < 0.95:
            low_conf.append({
                "jul_page": abs_jul_page,
                "matched_yabout_page": y_idx0_excl_first + 2,
                "score": round(score, 4),
            })
        prev_y_idx_excl_first = y_idx0_excl_first + 1

    for k in range(prev_y_idx_excl_first, len(yabout_infos) - 1):
        abs_page_1based = k + 2
        plan.append(PageRef("yabout", abs_page_1based, "inserted_yabout_tail"))

    inserted_yabout_pages_abs = [p.page_no for p in plan if p.source == "yabout"]
    jul_pages_in_output = [p.page_no for p in plan if p.source == "jul"]

    verification_checks = [
        {
            "name": "줄 파일 모든 페이지가 정확히 1번씩 유지되는가",
            "passed": jul_pages_in_output == list(range(1, len(jul_infos) + 1)),
            "actual": jul_pages_in_output,
            "expected": list(range(1, len(jul_infos) + 1)),
        },
        {
            "name": "맨 앞 첫 추가 페이지가 야붙 1페이지인가",
            "passed": len(plan) >= 1 and plan[0].source == "yabout" and plan[0].page_no == 1,
            "actual": asdict(plan[0]) if plan else None,
            "expected": {"source": "yabout", "page_no": 1, "reason": "inserted_yabout_front"},
        },
        {
            "name": "줄 페이지는 모두 원본 줄 PDF를 소스로 사용했는가",
            "passed": all(p.source == "jul" for p in plan if p.reason == "original_jul"),
            "actual": [asdict(p) for p in plan if p.reason == "original_jul"][:8],
            "expected": "all original_jul pages must come from jul source",
        },
        {
            "name": "출력 총 페이지 수가 줄+삽입야붙 수와 일치하는가",
            "passed": len(plan) == len(jul_infos) + len(inserted_yabout_pages_abs),
            "actual": len(plan),
            "expected": len(jul_infos) + len(inserted_yabout_pages_abs),
        },
    ]

    verification = {
        "passed": all(c["passed"] for c in verification_checks),
        "checks": verification_checks,
    }

    report = {
        "jul_file": str(jul_path),
        "yabout_file": str(yabout_path),
        "page_counts": {
            "jul": len(jul_infos),
            "yabout": len(yabout_infos),
            "logical_output": len(plan),
        },
        "inserted_yabout_pages": inserted_yabout_pages_abs,
        "matching_summary": {
            "total_matches": len(matches),
            "low_confidence_matches_below_0.95": low_conf,
        },
        "page_plan": [asdict(p) for p in plan],
        "verification": verification,
        "note_preservation_rule": "줄 페이지는 최종 생성 시 반드시 원본 줄 PDF에서 직접 렌더링",
        "merge_pdf_policy": "중간병합본도 원본 페이지를 직접 insert_pdf 하여 벡터/폰트/필기를 보존",
    }
    return plan, report


def write_merged_pdf_direct(page_plan: List[PageRef], source_pdfs: Dict[str, Path], output_pdf: Path) -> None:
    """Create the intermediate merged PDF by directly copying original PDF pages.
    This preserves original vectors, embedded fonts, and handwritten page content.
    """
    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    docs = {name: fitz.open(path) for name, path in source_pdfs.items()}
    out = fitz.open()
    try:
        for pref in page_plan:
            src_doc = docs[pref.source]
            page_index = pref.page_no - 1
            out.insert_pdf(src_doc, from_page=page_index, to_page=page_index)
        out.save(output_pdf)
    finally:
        out.close()
        for doc in docs.values():
            doc.close()


def write_merge_outputs(jul_path: Path, yabout_path: Path, out_dir: Path, page_plan: List[PageRef], report: dict) -> Tuple[Path, Path, Path, Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    output_stem = build_output_stem(jul_path)
    output_pdf = out_dir / f"{output_stem}.pdf"
    report_json = out_dir / f"{output_stem}_검증보고서.json"
    report_txt = out_dir / f"{output_stem}_검증요약.txt"
    page_plan_json = out_dir / f"{output_stem}_페이지계획.json"
    output_zip = out_dir / f"{output_stem}.zip"

    write_merged_pdf_direct(page_plan, {"jul": jul_path, "yabout": yabout_path}, output_pdf)

    report_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    page_plan_json.write_text(json.dumps([asdict(p) for p in page_plan], ensure_ascii=False, indent=2), encoding="utf-8")

    lines = []
    lines.append(f"줄 파일: {jul_path.name}")
    lines.append(f"야붙 파일: {yabout_path.name}")
    lines.append(f"중간병합본: {output_pdf.name}")
    lines.append(f"출력 총 페이지 수: {report['page_counts']['logical_output']}")
    lines.append(f"삽입된 야붙 페이지: {report['inserted_yabout_pages']}")
    lines.append(f"검증 통과: {report['verification']['passed']}")
    lines.append("")
    for check in report["verification"]["checks"]:
        lines.append(f"- {check['name']}: {'PASS' if check['passed'] else 'FAIL'}")
    lines.append("")
    lines.append("중간병합본 생성 방식: 원본 페이지 direct insert_pdf (벡터/폰트/필기 보존)")
    if report["matching_summary"]["low_confidence_matches_below_0.95"]:
        lines.append("낮은 확신도 매칭:")
        for item in report["matching_summary"]["low_confidence_matches_below_0.95"]:
            lines.append(f"  줄 {item['jul_page']}p <-> 야붙 {item['matched_yabout_page']}p (score={item['score']})")
    report_txt.write_text("\n".join(lines), encoding="utf-8")

    with zipfile.ZipFile(output_zip, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for p in [output_pdf, report_json, report_txt, page_plan_json]:
            zf.write(p, arcname=p.name)

    return output_pdf, output_zip, report_json, page_plan_json


def main() -> None:
    parser = argparse.ArgumentParser(description="줄 PDF에 야붙 페이지를 추가한 중간병합본을 생성")
    parser.add_argument("--jul", type=str, help="줄 PDF 경로")
    parser.add_argument("--yabout", type=str, help="야붙 PDF 경로")
    parser.add_argument("--folder", type=str, default=None, help="자동 탐지할 폴더")
    parser.add_argument("--outdir", type=str, default="output_merge", help="출력 폴더")
    args = parser.parse_args()

    if args.jul and args.yabout:
        jul_path = Path(args.jul).expanduser().resolve()
        yabout_path = Path(args.yabout).expanduser().resolve()
    else:
        folder = Path(args.folder or ".").expanduser().resolve()
        jul_path, yabout_path = auto_detect(folder)

    out_dir = Path(args.outdir).expanduser().resolve()
    page_plan, report = compute_page_plan(jul_path, yabout_path)
    output_pdf, output_zip, report_json, page_plan_json = write_merge_outputs(jul_path, yabout_path, out_dir, page_plan, report)
    print(output_pdf)
    print(output_zip)
    print(report_json)
    print(page_plan_json)


if __name__ == "__main__":
    main()

