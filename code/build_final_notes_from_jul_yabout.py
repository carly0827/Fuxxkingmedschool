from __future__ import annotations

import argparse
import json
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

import fitz  # PyMuPDF

from merge_jul_and_yabout import auto_detect, compute_page_plan, write_merge_outputs
from transcript_note_generator import generate_annotated_pdf_from_plan


import unicodedata


def _nfc(s: str) -> str:
    return unicodedata.normalize('NFC', s)


def ascii_safe_name(name: str) -> str:
    import re
    base = unicodedata.normalize('NFKD', name).encode('ascii', 'ignore').decode('ascii')
    base = re.sub(r'[^A-Za-z0-9._-]+', '_', base).strip('._')
    return base or 'artifact'


def build_final_stem(jul_path: Path) -> str:
    stem = _nfc(jul_path.stem).rstrip()
    if stem.endswith('줄'):
        return stem[:-1] + '최종본'
    return stem + '_최종본'


def read_transcript_text(args) -> str:
    if args.transcript_file:
        return Path(args.transcript_file).expanduser().read_text(encoding='utf-8')
    if args.transcript_text:
        return args.transcript_text
    if args.transcript_stdin:
        return sys.stdin.read()
    if sys.stdin and not sys.stdin.isatty():
        data = sys.stdin.read()
        if data.strip():
            return data
    raise RuntimeError('전사문이 없습니다. --transcript-file, --transcript-text, --transcript-stdin 중 하나를 사용하세요.')


def rename_copy(src: Path, dst: Path) -> Path:
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)
    return dst


def _is_ascii_filename(name: str) -> bool:
    try:
        name.encode('ascii')
        return True
    except UnicodeEncodeError:
        return False


def _sandbox_link_for(path: Path) -> str | None:
    try:
        resolved = path.resolve()
    except Exception:
        resolved = path
    s = str(resolved)
    if s.startswith('/mnt/data/'):
        return 'sandbox:' + s
    return None


def _validate_zip_file(path: Path) -> dict:
    result = {'can_open': False, 'testzip': None, 'error': None}
    try:
        with zipfile.ZipFile(path, 'r') as zf:
            result['can_open'] = True
            result['testzip'] = zf.testzip()
    except Exception as e:
        result['error'] = f'{type(e).__name__}: {e}'
    return result


def _validate_pdf_file(path: Path) -> dict:
    result = {'can_open': False, 'page_count': 0, 'error': None}
    try:
        doc = fitz.open(path)
        try:
            result['can_open'] = True
            result['page_count'] = len(doc)
        finally:
            doc.close()
    except Exception as e:
        result['error'] = f'{type(e).__name__}: {e}'
    return result


def _validate_docx_file(path: Path) -> dict:
    result = {'can_open': False, 'is_zip_like': False, 'error': None}
    try:
        with zipfile.ZipFile(path, 'r') as zf:
            result['can_open'] = True
            result['is_zip_like'] = 'word/document.xml' in zf.namelist()
    except Exception as e:
        result['error'] = f'{type(e).__name__}: {e}'
    return result


def build_direct_link_safe_copies(named_artifacts: dict[str, str | None], outdir: Path, safe_stem: str) -> dict:
    alias_specs = {
        'final_pdf': f'{safe_stem}_final.pdf',
        'transcript_digest_pdf': f'{safe_stem}_summary.pdf',
        'transcript_digest_docx': f'{safe_stem}_summary.docx',
        'final_zip': f'{safe_stem}_bundle.zip',
        'safe_download_zip': f'{safe_stem}_safe_download.zip',
        'safe_full_archive_zip': f'{safe_stem}_safe_full_archive.zip',
    }
    created = {}
    for key, alias_name in alias_specs.items():
        src = named_artifacts.get(key)
        if not src:
            continue
        src_path = Path(src)
        if not src_path.exists():
            continue
        dst = outdir / alias_name
        try:
            same = src_path.resolve() == dst.resolve()
        except Exception:
            same = src_path == dst
        if not same:
            shutil.copy2(src_path, dst)
        created[key] = str(dst if not same else src_path)
    return created


def verify_delivery_artifacts(named_artifacts: dict[str, str | None], outdir: Path, safe_stem: str) -> tuple[Path, Path, dict]:
    direct_safe = build_direct_link_safe_copies(named_artifacts, outdir, safe_stem)
    preferred_keys = ['final_pdf', 'transcript_digest_pdf', 'transcript_digest_docx', 'safe_download_zip', 'safe_full_archive_zip', 'final_zip']

    checks = []
    for key in preferred_keys:
        orig = named_artifacts.get(key)
        safe_copy = direct_safe.get(key)
        candidate = Path(safe_copy) if safe_copy else (Path(orig) if orig else None)
        entry = {
            'key': key,
            'original_path': orig,
            'link_target_path': str(candidate) if candidate else None,
            'link_target_is_direct_safe_copy': bool(safe_copy),
            'exists': False,
            'size_bytes': 0,
            'filename_ascii_only': None,
            'sandbox_link_candidate': None,
            'can_open': None,
            'details': {},
            'passed': False,
            'warning': None,
        }
        if candidate and candidate.exists():
            entry['exists'] = True
            entry['size_bytes'] = candidate.stat().st_size
            entry['filename_ascii_only'] = _is_ascii_filename(candidate.name)
            entry['sandbox_link_candidate'] = _sandbox_link_for(candidate)
            suffix = candidate.suffix.lower()
            if suffix == '.pdf':
                details = _validate_pdf_file(candidate)
                entry['details'] = details
                entry['can_open'] = details['can_open'] and details['page_count'] > 0
            elif suffix == '.zip':
                details = _validate_zip_file(candidate)
                entry['details'] = details
                entry['can_open'] = details['can_open'] and not details['testzip']
            elif suffix == '.docx':
                details = _validate_docx_file(candidate)
                entry['details'] = details
                entry['can_open'] = details['can_open'] and details['is_zip_like']
            else:
                entry['can_open'] = True
            entry['passed'] = bool(entry['exists'] and entry['size_bytes'] > 0 and entry['filename_ascii_only'] and entry['can_open'])
            if orig and (('%20' in orig) or any(ch in orig for ch in [' ', '(', ')'])):
                entry['warning'] = '원본 경로는 공백/기호를 포함할 수 있어 직접 링크 대신 ASCII 안전 사본 링크 사용 권장'
        else:
            entry['warning'] = '링크 대상 파일이 존재하지 않음'
        checks.append(entry)

    summary = {
        'verification_stage': '전달 전 링크/열기 검증',
        'rule': '사용자에게 링크를 보내기 전에 direct safe copy를 우선 사용하고, exists/size/open/ascii 여부를 모두 통과해야 함',
        'all_passed': all(c['passed'] for c in checks if c['key'] in {'final_pdf', 'safe_download_zip'}),
        'checks': checks,
        'preferred_direct_link_order': [c['sandbox_link_candidate'] for c in checks if c['passed'] and c['sandbox_link_candidate']],
        'fallback_instruction': '직접 PDF 링크가 불안정하면 safe_download_zip 링크를 먼저 전달',
    }

    json_path = outdir / f'{safe_stem}_delivery_link_check.json'
    txt_path = outdir / f'{safe_stem}_delivery_link_check.txt'
    json_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')

    lines = [
        '[전달 전 링크/열기 검증 요약]',
        f"all_passed={summary['all_passed']}",
        '핵심 규칙: 사용자에게는 원본 한글/공백 경로보다 direct safe copy 링크를 우선 전달할 것.',
        '',
    ]
    for c in checks:
        lines.extend([
            f"- {c['key']}",
            f"  passed={c['passed']}",
            f"  target={c['link_target_path']}",
            f"  sandbox={c['sandbox_link_candidate']}",
            f"  size_bytes={c['size_bytes']}",
            f"  filename_ascii_only={c['filename_ascii_only']}",
            f"  can_open={c['can_open']}",
        ])
        if c['warning']:
            lines.append(f"  warning={c['warning']}")
        if c['details']:
            lines.append(f"  details={json.dumps(c['details'], ensure_ascii=False)}")
        lines.append('')
    txt_path.write_text('\n'.join(lines), encoding='utf-8')
    return json_path, txt_path, summary


def build_safe_aliases(artifacts, safe_dir: Path) -> list[tuple[str, Path]]:
    safe_dir.mkdir(parents=True, exist_ok=True)
    aliases: list[tuple[str, Path]] = []
    for idx, (label, p) in enumerate(artifacts, start=1):
        if not p or not Path(p).exists():
            continue
        suffix = Path(p).suffix or '.bin'
        safe_name = f"{idx:02d}_{label}{suffix}"
        dst = safe_dir / safe_name
        shutil.copy2(p, dst)
        aliases.append((label, dst))
    info_lines = [
        'Open/download-safe aliases for iOS and compressed delivery.',
        'Use the safe_download zip first when the direct PDF or Korean filename does not open well.',
    ]
    (safe_dir / '00_README_SAFE_DOWNLOAD.txt').write_text('\n'.join(info_lines), encoding='utf-8')
    return aliases


def run_pipeline(jul_path: Path, yabout_path: Path, transcript_text: str, outdir: Path, keep_merged_pdf: bool = True) -> dict:
    outdir.mkdir(parents=True, exist_ok=True)
    final_stem = build_final_stem(jul_path)

    with tempfile.TemporaryDirectory(prefix='jul_yabout_note_v4_') as tmp:
        tmpdir = Path(tmp)
        merge_dir = tmpdir / 'merge_stage'
        note_dir = tmpdir / 'note_stage'

        page_plan, merge_report = compute_page_plan(jul_path, yabout_path)
        merged_pdf, merged_zip, merge_report_json, page_plan_json = write_merge_outputs(
            jul_path, yabout_path, merge_dir, page_plan, merge_report
        )
        merge_summary_txt = merged_pdf.with_name(merged_pdf.stem + '_검증요약.txt')

        annotated_pdf, precheck_json, report_json, completeness_json, completeness_txt = generate_annotated_pdf_from_plan(
            page_plan=merge_report['page_plan'],
            source_pdf_map={'jul': str(jul_path), 'yabout': str(yabout_path)},
            transcript_text=transcript_text,
            output_dir=str(note_dir),
        )
        digest_pdf = note_dir / 'overall_transcript_summary_digest.pdf'
        digest_docx = note_dir / 'overall_transcript_summary_digest.docx'
        digest_generation_json = note_dir / 'overall_transcript_summary_digest_generation.json'

        final_pdf = rename_copy(Path(annotated_pdf), outdir / f'{final_stem}.pdf')
        final_precheck = rename_copy(Path(precheck_json), outdir / f'{final_stem}_전사문사전검증.json')
        final_report = rename_copy(Path(report_json), outdir / f'{final_stem}_전사문검증보고서.json')
        final_completeness_json = rename_copy(Path(completeness_json), outdir / f'{final_stem}_전사문누락검토.json')
        final_completeness_txt = rename_copy(Path(completeness_txt), outdir / f'{final_stem}_전사문누락검토.txt')
        final_digest_pdf = rename_copy(digest_pdf, outdir / f'{final_stem}_전체전사문요약정리표.pdf') if digest_pdf.exists() else None
        final_digest_docx = rename_copy(digest_docx, outdir / f'{final_stem}_전체전사문요약정리표.docx') if digest_docx.exists() else None
        final_digest_generation_json = rename_copy(digest_generation_json, outdir / f'{final_stem}_전체전사문요약정리표_생성정보.json') if digest_generation_json.exists() else None
        final_merge_report = rename_copy(Path(merge_report_json), outdir / f'{final_stem}_줄야붙병합검증.json')
        final_merge_summary = rename_copy(Path(merge_summary_txt), outdir / f'{final_stem}_줄야붙병합요약.txt')
        final_page_plan = rename_copy(Path(page_plan_json), outdir / f'{final_stem}_페이지계획.json')
        final_merged_pdf = None
        final_merged_zip = None
        if keep_merged_pdf:
            final_merged_pdf = rename_copy(Path(merged_pdf), outdir / f'{final_stem}_중간병합본.pdf')
        if Path(merged_zip).exists():
            final_merged_zip = rename_copy(Path(merged_zip), outdir / f'{final_stem}_중간병합본.zip')

    manifest = {
        'jul_file': str(jul_path),
        'yabout_file': str(yabout_path),
        'final_pdf': str(final_pdf),
        'transcript_precheck_json': str(final_precheck),
        'transcript_report_json': str(final_report),
        'transcript_completeness_json': str(final_completeness_json),
        'transcript_completeness_txt': str(final_completeness_txt),
        'transcript_digest_pdf': str(final_digest_pdf) if final_digest_pdf else None,
        'transcript_digest_docx': str(final_digest_docx) if final_digest_docx else None,
        'transcript_digest_generation_json': str(final_digest_generation_json) if final_digest_generation_json else None,
        'merge_report_json': str(final_merge_report),
        'merge_summary_txt': str(final_merge_summary),
        'page_plan_json': str(final_page_plan),
        'merged_pdf_intermediate': str(final_merged_pdf) if final_merged_pdf else None,
        'merged_zip_intermediate': str(final_merged_zip) if final_merged_zip else None,
        'transcript_char_count': len(transcript_text),
        'naming_rule': '원래 줄 파일 이름의 마지막 줄 -> 최종본',
        'critical_preservation_rule': '줄 파일 페이지는 최종본 생성 시 원본 줄 PDF에서 직접 렌더링하며, 야붙의 대응 페이지로 대체하지 않음',
        'merged_pdf_preservation_rule': '중간병합본도 원본 페이지를 직접 insert_pdf 하므로 폰트/벡터/필기 내용이 깨지지 않아야 함',
        'continuation_rule': '전사문은 길이에 상관없이 모든 내용을 유지하고 필요한 만큼 continuation page 생성',
        'digest_generation_rule': '전체 전사문 요약정리표는 DOCX를 먼저 생성하고, 그 DOCX를 PDF로 변환한다. 변환 실패 시에만 이미지 기반 PDF를 보조적으로 사용한다.',
        'delivery_recommendation': '직접 PDF 대신 safe_download zip을 우선 전달하도록 권장',
    }

    manifest_path = outdir / f'{final_stem}_생성과정요약.json'
    final_zip = outdir / f'{final_stem}.zip'
    safe_zip = outdir / f'{ascii_safe_name(final_stem)}_safe_download.zip'
    safe_full_zip = outdir / f'{ascii_safe_name(final_stem)}_safe_full_archive.zip'
    manifest['final_zip'] = str(final_zip)
    manifest['safe_download_zip'] = str(safe_zip)
    manifest['safe_full_archive_zip'] = str(safe_full_zip)
    manifest['preferred_delivery_file'] = str(safe_zip)
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')

    artifacts = [
        ("final_pdf", final_pdf),
        ("transcript_precheck", final_precheck),
        ("transcript_report", final_report),
        ("transcript_completeness_json", final_completeness_json),
        ("transcript_completeness_txt", final_completeness_txt),
        ("transcript_digest_pdf", final_digest_pdf),
        ("transcript_digest_docx", final_digest_docx),
        ("transcript_digest_generation_json", final_digest_generation_json),
        ("merge_report_json", final_merge_report),
        ("merge_summary_txt", final_merge_summary),
        ("page_plan_json", final_page_plan),
        ("manifest_json", manifest_path),
    ]
    if final_merged_pdf and Path(final_merged_pdf).exists():
        artifacts.append(("merged_pdf_intermediate", final_merged_pdf))
    if final_merged_zip and Path(final_merged_zip).exists():
        artifacts.append(("merged_zip_intermediate", final_merged_zip))

    with zipfile.ZipFile(final_zip, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
        for _, p in artifacts:
            if p and Path(p).exists():
                zf.write(p, arcname=Path(p).name)

    core_labels = {'final_pdf', 'transcript_digest_pdf', 'transcript_digest_docx', 'transcript_precheck', 'transcript_report', 'transcript_completeness_json', 'transcript_completeness_txt', 'manifest_json'}
    core_artifacts = [(label, p) for label, p in artifacts if label in core_labels]
    safe_alias_dir = outdir / f'{ascii_safe_name(final_stem)}_safe_alias'
    safe_aliases = build_safe_aliases(core_artifacts, safe_alias_dir)

    with zipfile.ZipFile(safe_zip, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
        readme_path = safe_alias_dir / '00_README_SAFE_DOWNLOAD.txt'
        if readme_path.exists():
            zf.write(readme_path, arcname=readme_path.name)
        for _, p in safe_aliases:
            zf.write(p, arcname=Path(p).name)

    full_safe_alias_dir = outdir / f'{ascii_safe_name(final_stem)}_safe_alias_full'
    full_safe_aliases = build_safe_aliases(artifacts, full_safe_alias_dir)
    with zipfile.ZipFile(safe_full_zip, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
        readme_path = full_safe_alias_dir / '00_README_SAFE_DOWNLOAD.txt'
        if readme_path.exists():
            zf.write(readme_path, arcname=readme_path.name)
        for _, p in full_safe_aliases:
            zf.write(p, arcname=Path(p).name)

    named_artifacts = {
        'final_pdf': str(final_pdf),
        'transcript_digest_pdf': str(final_digest_pdf) if final_digest_pdf else None,
        'transcript_digest_docx': str(final_digest_docx) if final_digest_docx else None,
        'final_zip': str(final_zip),
        'safe_download_zip': str(safe_zip),
        'safe_full_archive_zip': str(safe_full_zip),
    }
    delivery_link_check_json, delivery_link_check_txt, delivery_link_summary = verify_delivery_artifacts(
        named_artifacts=named_artifacts,
        outdir=outdir,
        safe_stem=ascii_safe_name(final_stem),
    )
    manifest['delivery_link_check_json'] = str(delivery_link_check_json)
    manifest['delivery_link_check_txt'] = str(delivery_link_check_txt)
    manifest['delivery_link_verification_rule'] = '사용자에게 링크를 보낼 때는 delivery_link_check에서 passed=true인 direct safe copy를 우선 사용'
    manifest['delivery_preferred_direct_links'] = delivery_link_summary.get('preferred_direct_link_order', [])
    manifest['delivery_all_passed'] = delivery_link_summary.get('all_passed', False)

    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
    return manifest


def main() -> None:
    parser = argparse.ArgumentParser(description='줄 PDF + 야붙 PDF + 전사문으로 최종 필기본 PDF를 생성하는 통합 파이프라인')
    parser.add_argument('--jul', type=str, help="파일명에 '줄'이 포함된 원본/필기 슬라이드 PDF")
    parser.add_argument('--yabout', type=str, help="파일명에 '야붙'이 포함된 추가문제/모아보기 PDF")
    parser.add_argument('--folder', type=str, default=None, help='자동 탐지할 폴더')
    parser.add_argument('--transcript-file', type=str, default=None, help='전사문 txt 파일 경로')
    parser.add_argument('--transcript-text', type=str, default=None, help='전사문 전체 텍스트를 직접 전달')
    parser.add_argument('--transcript-stdin', action='store_true', help='표준입력(stdin)에서 전사문 읽기')
    parser.add_argument('--outdir', type=str, default='output_final_notes', help='출력 폴더')
    parser.add_argument('--no-keep-merged-pdf', action='store_true', help='중간 병합 PDF를 최종 결과 폴더에 복사하지 않음')
    args = parser.parse_args()

    if args.jul and args.yabout:
        jul_path = Path(args.jul).expanduser().resolve()
        yabout_path = Path(args.yabout).expanduser().resolve()
    else:
        folder = Path(args.folder or '.').expanduser().resolve()
        jul_path, yabout_path = auto_detect(folder)

    if not jul_path.exists():
        raise FileNotFoundError(f'줄 파일을 찾을 수 없습니다: {jul_path}')
    if not yabout_path.exists():
        raise FileNotFoundError(f'야붙 파일을 찾을 수 없습니다: {yabout_path}')

    transcript_text = read_transcript_text(args)
    if not transcript_text.strip():
        raise RuntimeError('전사문 텍스트가 비어 있습니다.')

    outdir = Path(args.outdir).expanduser().resolve()
    manifest = run_pipeline(
        jul_path=jul_path,
        yabout_path=yabout_path,
        transcript_text=transcript_text,
        outdir=outdir,
        keep_merged_pdf=not args.no_keep_merged_pdf,
    )

    print('완료')
    print(f"FINAL PDF: {manifest['final_pdf']}")
    print(f"FINAL ZIP: {manifest['final_zip']}")
    print(f"SAFE ZIP: {manifest['safe_download_zip']}")
    print(f"SAFE FULL ZIP: {manifest['safe_full_archive_zip']}")
    print(f"PREFERRED DELIVERY: {manifest['preferred_delivery_file']}")
    print(f"TRANSCRIPT COMPLETENESS: {manifest['transcript_completeness_json']}")
    print(f"PAGE PLAN: {manifest['page_plan_json']}")
    print(f"MANIFEST: {outdir / (build_final_stem(jul_path) + '_생성과정요약.json')}")


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(f'오류: {e}', file=sys.stderr)
        sys.exit(1)

