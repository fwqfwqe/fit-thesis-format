#!/usr/bin/env python3
"""Recheck a final DOCX and its evidence. A successful command is not a pass."""
import argparse
import hashlib
import json
import re
from pathlib import Path
import subprocess
import sys
import tempfile

from format_audit import Audit
from thesis_format import package, inventory, compare_content

SCRIPTS = Path(__file__).resolve().parent
PAGE_CHECKS = ('layout', 'headers_footers', 'numbering_toc', 'math_tables', 'content_order')


def continuation_requirements(rendering, final):
    from thesis_format import parse, q, PROTECTED_TAGS
    root = parse(package(final)['word/document.xml'])
    table_count = sum(1 for t in root.iter(q('tbl')) if not any(a.tag in PROTECTED_TAGS for a in t.iterancestors()))
    spans = rendering.get('table_spans', [])
    if not isinstance(spans, list) or len(spans) != table_count:
        raise ValueError('Native table/page inventory does not cover all document tables')
    required, seen = {}, set()
    count = rendering.get('pages', 0)
    for item in spans:
        index, first, last = (item.get(k) for k in ('table_index','start_page','end_page'))
        if type(index) is not int or index in seen or not 1 <= index <= table_count:
            raise ValueError('Native table identifiers are missing or duplicated')
        if type(first) is not int or type(last) is not int or not 1 <= first <= last <= count:
            raise ValueError('Native table page span is invalid')
        seen.add(index)
        for number in range(first+1, last+1):required.setdefault(number, []).append(index)
    return required


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def read(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))


def evidence_path(value, parent):
    if not isinstance(value, str) or not value.strip():
        raise ValueError('Missing evidence file path')
    path = Path(value).expanduser()
    return (path if path.is_absolute() else parent / path).resolve()


def verify_chain(source, final, chain_path):
    """Replay existing content-preserving verifiers, never trust saved ok flags."""
    chain = read(chain_path)
    if chain.get('original_sha256') != sha(source):
        raise ValueError('Chain is not bound to the original source')
    stages = chain.get('stages')
    if not isinstance(stages, list) or not stages:
        raise ValueError('Chain must contain verified stages')
    previous = sha(source)
    summaries = []
    with tempfile.TemporaryDirectory(prefix='fit-delivery-check-') as temp:
        for index, stage in enumerate(stages):
            if isinstance(stage, dict) and stage.get('kind') in ('format', 'office_fields'):
                from integrity import verify_transition
                input_path = evidence_path(stage.get('source'), chain_path.parent)
                output = evidence_path(stage.get('output'), chain_path.parent)
                report_path = evidence_path(stage.get('report'), chain_path.parent)
                if sha(input_path) != previous or stage.get('output_sha256') != sha(output):
                    raise ValueError('Disconnected or changed registered stage')
                result = verify_transition(input_path, output, stage['kind'], report=read(report_path))
                if result.get('passed') is not True:
                    raise ValueError('Registered stage failed protection: ' + str(result.get('errors')))
                stage_report = read(report_path)
                source_map = stage_report.get('source_mapping')
                if source_map is None:
                    source_map = next((op.get('source_mapping') for op in stage_report.get('operations', [])
                                       if op.get('kind') == 'format_roles'), None)
                summaries.append({'scope': stage['kind'], 'source_sha256': sha(input_path),
                                  'output_sha256': sha(output), 'report_sha256': sha(report_path),
                                  'source_mapping': source_map})
                previous = sha(output)
                continue
            if not isinstance(stage, dict) or stage.get('scope') not in ('styles', 'layout', 'citations', 'narrative_citations'):
                raise ValueError('Unsupported stage; custom structural edits need a separately implemented replay verifier')
            input_path = evidence_path(stage.get('source'), chain_path.parent)
            output = evidence_path(stage.get('output'), chain_path.parent)
            mapping = evidence_path(stage.get('mapping'), chain_path.parent)
            if sha(input_path) != previous:
                raise ValueError(f'Stage {index + 1} is disconnected from the original chain')
            request = {'action': 'verify', 'scope': stage['scope'], 'source': str(input_path),
                       'output': str(output), 'mapping': str(mapping),
                       'report': str(Path(temp) / f'{index}.json')}
            run = subprocess.run([sys.executable, str(SCRIPTS / 'agent_cli.py'), '--request', '-'],
                                 input=json.dumps(request), text=True, encoding='utf-8', capture_output=True)
            result = json.loads(run.stdout)
            if run.returncode != 0 or result.get('ok') is not True:
                raise ValueError(f'Stage {index + 1} did not pass a fresh verifier: ' + run.stdout[-1500:])
            summaries.append({'scope': stage['scope'], 'source_sha256': sha(input_path),
                              'output_sha256': sha(output), 'mapping_sha256': sha(mapping)})
            previous = sha(output)
    if previous != sha(final):
        raise ValueError('Final DOCX differs from the last verified stage')
    return summaries


def check_render(final, render_path, pages_path, review_path, allow_incomplete=False):
    if not all((render_path, pages_path, review_path)):
        raise ValueError('Need the real export report, all-page image manifest, and completed page review')
    rendering, pages, review = read(render_path), read(pages_path), read(review_path)
    final_hash = sha(final)
    rendered_hash = rendering.get('saved_docx_sha256') or rendering.get('source_sha256_before', '')
    if not isinstance(rendered_hash, str) or rendered_hash.lower() != final_hash:
        raise ValueError('Rendered DOCX is not this final DOCX')
    if rendering.get('source_unchanged') is not True:
        raise ValueError('Export did not verify source preservation')
    if not rendering.get('engine') or not rendering.get('engine_version'):
        raise ValueError('Export must identify the tested document engine and version')
    pdf = evidence_path(rendering.get('pdf'), render_path.parent)
    pdf_hash = sha(pdf)
    if rendering.get('pdf_sha256', '').lower() != pdf_hash:
        raise ValueError('PDF does not match the export evidence')
    if pages.get('input_sha256', '').lower() != pdf_hash:
        raise ValueError('Page images were made from a different PDF')
    if review.get('final_docx_sha256', '').lower() != final_hash or review.get('pdf_sha256', '').lower() != pdf_hash:
        raise ValueError('Page review belongs to a different document or PDF')
    if not isinstance(review.get('reviewer'), str) or not review['reviewer'].strip():
        raise ValueError('Page review must name the actual reviewing agent/person')
    count = rendering.get('pages')
    if type(count) is not int or count < 1 or pages.get('page_count') != count:
        raise ValueError('Page counts are missing or inconsistent')
    dpi = pages.get('dpi')
    if pages.get('engine') != 'pypdfium2' or type(dpi) is not int or not 72 <= dpi <= 300:
        raise ValueError('Page proof requires a reproducible pypdfium2 rasterization and 72–300 DPI')
    try:
        import pypdfium2 as pdfium
        from PIL import Image
    except ImportError as exc:
        raise ValueError('PDF page count must be verified independently; install requirements-render.txt for the final delivery check') from exc
    with pdfium.PdfDocument(str(pdf)) as actual_pdf:
        if len(actual_pdf) != count:
            raise ValueError('Submitted page count differs from the actual PDF; some pages are missing from review')
    image_items, review_items = pages.get('pages'), review.get('pages')
    if not isinstance(image_items, list) or not isinstance(review_items, list):
        raise ValueError('Page image/review lists are missing')
    image_map = {item['page']: item for item in image_items}
    review_map = {item['page']: item for item in review_items}
    expected = set(range(1, count + 1))
    if len(image_items) != count or len(review_items) != count or set(image_map) != expected or set(review_map) != expected:
        raise ValueError('Every page must have exactly one image and review entry')
    continued = continuation_requirements(rendering, final)
    pending_pages, failed_pages = [], []
    for number in sorted(expected):
        item, inspected = image_map[number], review_map[number]
        image_path = evidence_path(item.get('file'), pages_path.parent)
        image_hash = sha(image_path)
        with Image.open(image_path) as picture:
            if picture.format not in ('PNG', 'JPEG', 'WEBP'):
                raise ValueError(f'Page {number} must be a readable page image')
            picture.verify()
        if item.get('sha256', '').lower() != image_hash or inspected.get('image_sha256', '').lower() != image_hash:
            raise ValueError(f'Page {number} image differs from the reviewed image')
        # Matching submitted hashes cannot establish that a PNG depicts this
        # PDF page. Re-render its pixels independently from the actual PDF.
        with pdfium.PdfDocument(str(pdf)) as check_pdf:
            page = check_pdf[number - 1]
            bitmap = page.render(scale=dpi / 72)
            regenerated = bitmap.to_pil()
            try:
                with Image.open(image_path) as submitted:
                    if submitted.size != regenerated.size or submitted.convert('RGB').tobytes() != regenerated.convert('RGB').tobytes():
                        raise ValueError(f'Page {number} image pixels do not reproduce the actual PDF page')
            finally:
                regenerated.close(); bitmap.close(); page.close()
        status = inspected.get('status')
        if allow_incomplete and status in ('pending', 'fail'):
            if not isinstance(inspected.get('checks'), dict) or any(inspected['checks'].get(check) not in ('pass','fail','pending') for check in PAGE_CHECKS):
                raise ValueError(f'Page {number} has invalid review states')
            if status == 'fail' and not str(inspected.get('notes', '')).strip():
                raise ValueError(f'Page {number} failure needs concrete observations')
            (pending_pages if status == 'pending' else failed_pages).append(number)
            continue
        if status != 'pass' or not isinstance(inspected.get('checks'), dict):
            raise ValueError(f'Page {number} was not recorded as checked and passed')
        if any(inspected['checks'].get(check) != 'pass' for check in PAGE_CHECKS):
            raise ValueError(f'Page {number} has an incomplete or failed checklist')
        labels = inspected.get('continued_tables', [])
        if not isinstance(labels, list):
            raise ValueError(f'Page {number} continued_tables must be a list')
        label_map = {entry.get('table_index'): entry for entry in labels}
        if len(label_map) != len(labels) or set(label_map) != set(continued.get(number, [])):
            raise ValueError(f'Page {number} must record exactly the native continued tables')
        for index, label in label_map.items():
            if not isinstance(label.get('caption_label'), str) or re.search(r'续前表|续表', label['caption_label']) is None:
                raise ValueError(f'Page {number} table {index} has no observed continuation caption')
            if label.get('header_repeated') is not True or label.get('no_missing_rows') is not True:
                raise ValueError(f'Page {number} table {index} has not passed repeated-header/row-integrity review')
        if not isinstance(inspected.get('notes'), str) or not inspected['notes'].strip():
            raise ValueError(f'Page {number} lacks concrete observations')
    return {'engine': rendering['engine'], 'engine_version': rendering['engine_version'],
            'page_count': count, 'pdf_sha256': pdf_hash, 'reviewer': review['reviewer'],
            'all_pages_passed': not pending_pages and not failed_pages,
            'pending_pages': pending_pages, 'failed_pages': failed_pages,
            'native_math_count': rendering.get('native_math_count', 0),
            'table_spans_verified': True,
            'continued_tables_checked': not pending_pages and not failed_pages,
            'continued_page_checks': sum(len(v) for v in continued.values()),
            'evidence_consistent': True, 'visual_inspection_performed_by_this_script': False,
            'page_pixels_recomputed': True,
            'note': 'Hashes bind the submitted evidence to the final files. The script cannot prove the reviewer actually looked at or correctly interpreted every page.'}


def check_dual_render(final, evidence):
    """Require two independently exported, fully reviewed views of the same file."""
    if not isinstance(evidence, dict) or set(evidence) != {'word', 'wps'}:
        raise ValueError('Both Word and WPS evidence are mandatory')
    results = {}
    from thesis_format import parse
    parts = package(final)
    pagerefs = parse(parts['word/document.xml']).xpath('//*[local-name()="instrText" and contains(text(),"PAGEREF")]')
    has_toc = bool(pagerefs)
    for engine in ('word', 'wps'):
        proof = evidence[engine]
        if not isinstance(proof, dict) or any(not proof.get(k) for k in ('render_manifest', 'page_manifest', 'review')):
            raise ValueError(engine + ': missing export, page manifest, or actual page review')
        paths = {key: Path(proof[key]).resolve() for key in ('render_manifest', 'page_manifest', 'review')}
        report = read(paths['render_manifest'])
        if report.get('engine_key') != engine:
            raise ValueError('Engine identity does not match the requested engine')
        executable = str(report.get('engine_executable', '')).lower().replace('\\', '/')
        if engine == 'wps' and not any(word in executable for word in ('kingsoft', 'wps', 'kwps')):
            raise ValueError('WPS must be identified by the actual application executable, not its Word-compatible name')
        if engine == 'word' and ('winword' not in executable or any(word in executable for word in ('kingsoft', 'wps'))):
            raise ValueError('Word executable identity is missing or inconsistent')
        if report.get('source_sha256_before', '').lower() != sha(final) or report.get('saved_docx'):
            raise ValueError('Dual verification must read the exact final DOCX, without exporting separate saved versions')
        if report.get('field_errors'):
            raise ValueError(engine + ': unresolved native field errors')
        entries = report.get('toc_entries')
        if has_toc and not entries:
            raise ValueError(engine + ': native TOC target/page evidence is missing')
        if has_toc and len(entries) < len(pagerefs):
            raise ValueError(engine + ': some TOC page references were not checked')
        for item in entries or []:
            if not item.get('target_bookmark') or item.get('passed') is not True:
                raise ValueError(engine + ': TOC cache does not match native target pages')
            if item.get('actual_page') is None or item.get('displayed_page') is None:
                raise ValueError(engine + ': incomplete native TOC page evidence')
            actual = int(item['actual_page'])
            def roman(value):
                text = ''
                for number, symbol in ((1000,'M'),(900,'CM'),(500,'D'),(400,'CD'),(100,'C'),(90,'XC'),(50,'L'),(40,'XL'),(10,'X'),(9,'IX'),(5,'V'),(4,'IV'),(1,'I')):
                    count, value = divmod(value, number)
                    text += symbol * count
                return text
            if str(item['displayed_page']).strip() not in {str(actual), roman(actual), roman(actual).lower()}:
                raise ValueError(engine + ': claimed TOC pass contradicts actual page numbers')
        results[engine] = check_render(final, paths['render_manifest'], paths['page_manifest'], paths['review'])
    results['same_final_docx_sha256'] = sha(final)
    return results


def check(source, final, chain=None, render_manifest=None, page_manifest=None, review=None, coverage=None,
          final_mapping=None, dual_render=None):
    before = {'source': sha(source), 'final': sha(final)}
    blockers, details = [], {}
    try:
        if chain:
            details['fresh_stage_verification'] = verify_chain(source, final, chain)
        else:
            if before['source'] != before['final']:
                raise ValueError('A changed DOCX requires a replayable stage chain. Text comparison alone cannot protect hyperlink targets, effective numbering or every package part.')
            errors = compare_content(inventory(package(source)), inventory(package(final)))
            if errors:
                raise ValueError('; '.join(errors) + '. Supply a valid supported stage chain; do not reset the original baseline.')
            details['content_verification'] = 'strict_original_content_unchanged'
    except Exception as exc:
        blockers.append({'check': 'content_preservation', 'reason': str(exc)})
    try:
        if coverage is None:
            raise ValueError('Whole-document review coverage is missing; use review_plan.py to review and account for every paragraph and table, including the end of the paper.')
        from review_plan import verify_coverage
        coverage_result = verify_coverage(source, coverage)
        details['whole_document_coverage'] = coverage_result
        mapping_path = Path(coverage_result['style_map_path'])
        mapping = read(mapping_path)
        mapped = bool(mapping.get('paragraphs') or mapping.get('tables') or any(mapping.get('story_paragraphs', {}).values()))
        if mapped:
            matched = any(stage['source_sha256'] == before['source'] and (
                          stage['scope'] == 'styles' and stage.get('mapping_sha256') == coverage_result['style_map_sha256']
                          or stage['scope'] == 'format' and stage.get('source_mapping') == mapping)
                          for stage in details.get('fresh_stage_verification', []))
            if not matched:
                raise ValueError('The complete reviewed style map was not applied in a freshly verified stage from the original source; a partial map is insufficient.')
    except Exception as exc:
        blockers.append({'check': 'whole_document_coverage', 'reason': str(exc)})
    try:
        if final_mapping is None:
            if before['source'] == before['final'] and coverage:
                final_map = read(Path(details['whole_document_coverage']['style_map_path']))
            else:
                raise ValueError('Final classification map is required; do not drop reviewed table roles during final audit')
        else:
            final_map = read(final_mapping)
        from thesis_job import verify_mapping_complete
        details['final_mapping_coverage'] = verify_mapping_complete(final, final_map)
        audit = Audit(final, mapping=final_map).run()
        unresolved = [item for item in audit['issues'] if item['severity'] in ('error', 'warning')]
        details['fresh_audit'] = {'source_sha256': audit['source_sha256'], 'summary': audit['summary'],
                                  'blocking_findings': unresolved}
        if unresolved:
            blockers.append({'check': 'format_requirements', 'reason': f'{len(unresolved)} unresolved error/warning findings in the final DOCX',
                             'categories': sorted({item['category'] for item in unresolved})})
    except Exception as exc:
        blockers.append({'check': 'readable_docx_and_static_audit', 'reason': str(exc)})
    try:
        if dual_render is None:
            raise ValueError('v10 requires Word and WPS evidence; supply --dual-render JSON containing both engines')
        details['render_and_page_evidence'] = check_dual_render(final, read(dual_render))
    except Exception as exc:
        blockers.append({'check': 'render_and_page_review', 'reason': str(exc)})
    if sha(source) != before['source'] or sha(final) != before['final']:
        blockers.append({'check': 'files_unchanged_during_check', 'reason': 'Input changed while being checked'})
    ready = not blockers
    return {'schema_version': 1, 'operation_succeeded': True, 'delivery_ready': ready,
            'status': 'checks_passed_with_review_record' if ready else 'needs_review',
            'source_sha256': before['source'], 'final_docx_sha256': before['final'],
            'blocking_items': blockers, 'details': details,
            'permitted_description': 'Checks passed with the stated page-review record' if ready else '未验收稿或部分完成；不得声称整篇排版合格',
            'limitations': ['This script checks files and evidence, not the honesty or accuracy of visual review.',
                            'Only registered transformations with fresh integrity verification are accepted; arbitrary allow-change flags are unsupported.',
                            'Detected warnings may require a focused correction to the auditor if they are false positives; do not erase findings or substitute a new source baseline.']}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--source', type=Path, required=True)
    parser.add_argument('--final', type=Path, required=True)
    parser.add_argument('--report', type=Path, required=True)
    for name in ('chain', 'render-manifest', 'page-manifest', 'review', 'coverage', 'final-mapping', 'dual-render'):
        parser.add_argument('--' + name, type=Path)
    args = parser.parse_args()
    try:
        inputs = [p.resolve() for p in (args.source, args.final, args.chain, args.render_manifest, args.page_manifest, args.review, args.coverage, args.final_mapping, args.dual_render) if p]
        report = args.report.expanduser().resolve()
        if report in inputs or report.exists() or report.is_symlink():
            raise ValueError('Report must be a new path distinct from every input')
        if args.source.suffix.lower() != '.docx' or args.final.suffix.lower() != '.docx':
            raise ValueError('Source and final must be DOCX files')
        result = check(args.source.resolve(), args.final.resolve(),
                       **{key: getattr(args, key).resolve() if getattr(args, key) else None
                          for key in ('chain', 'render_manifest', 'page_manifest', 'review', 'coverage', 'final_mapping', 'dual_render')})
        report.parent.mkdir(parents=True, exist_ok=True)
        with report.open('x', encoding='utf-8') as handle:
            json.dump(result, handle, ensure_ascii=False, indent=2)
            handle.write('\n')
        print(json.dumps({'operation_succeeded': True, 'delivery_ready': result['delivery_ready'],
                          'status': result['status'], 'report': str(report),
                          'blocking_items': result['blocking_items']}, ensure_ascii=False, indent=2))
        return 0 if result['delivery_ready'] else 3
    except Exception as exc:
        print(json.dumps({'operation_succeeded': False, 'delivery_ready': False, 'error': str(exc)}, ensure_ascii=False))
        return 2


if __name__ == '__main__':
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    raise SystemExit(main())
