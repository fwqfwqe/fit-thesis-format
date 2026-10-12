#!/usr/bin/env python3
"""Recheck a final DOCX and its evidence. A successful command is not a pass."""
import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile

from format_audit import Audit
from thesis_format import package, inventory, compare_content

SCRIPTS = Path(__file__).resolve().parent
PAGE_CHECKS = ('layout', 'headers_footers', 'numbering_toc', 'math_tables', 'content_order')


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


def check_render(final, render_path, pages_path, review_path):
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
        if inspected.get('status') != 'pass' or not isinstance(inspected.get('checks'), dict):
            raise ValueError(f'Page {number} was not recorded as checked and passed')
        if any(inspected['checks'].get(check) != 'pass' for check in PAGE_CHECKS):
            raise ValueError(f'Page {number} has an incomplete or failed checklist')
        if not isinstance(inspected.get('notes'), str) or not inspected['notes'].strip():
            raise ValueError(f'Page {number} lacks concrete observations')
    return {'engine': rendering['engine'], 'engine_version': rendering['engine_version'],
            'page_count': count, 'pdf_sha256': pdf_hash, 'reviewer': review['reviewer'],
            'evidence_consistent': True, 'visual_inspection_performed_by_this_script': False,
            'note': 'Hashes bind the submitted evidence to the final files. The script cannot prove the reviewer actually looked at or correctly interpreted every page.'}


def check(source, final, chain=None, render_manifest=None, page_manifest=None, review=None, coverage=None):
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
            matched = any(stage['scope'] == 'styles' and stage['source_sha256'] == before['source']
                          and stage['mapping_sha256'] == coverage_result['style_map_sha256']
                          for stage in details.get('fresh_stage_verification', []))
            if not matched:
                raise ValueError('The complete reviewed style map was not applied in a freshly verified stage from the original source; a partial map is insufficient.')
    except Exception as exc:
        blockers.append({'check': 'whole_document_coverage', 'reason': str(exc)})
    try:
        audit = Audit(final).run()
        unresolved = [item for item in audit['issues'] if item['severity'] in ('error', 'warning')]
        details['fresh_audit'] = {'source_sha256': audit['source_sha256'], 'summary': audit['summary'],
                                  'blocking_findings': unresolved}
        if unresolved:
            blockers.append({'check': 'format_requirements', 'reason': f'{len(unresolved)} unresolved error/warning findings in the final DOCX',
                             'categories': sorted({item['category'] for item in unresolved})})
    except Exception as exc:
        blockers.append({'check': 'readable_docx_and_static_audit', 'reason': str(exc)})
    try:
        details['render_and_page_evidence'] = check_render(final, render_manifest, page_manifest, review)
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
                            'Custom cover reconstruction, renumbering, equation transcription, appendix movement and field-cache refresh are not accepted via arbitrary allow-change flags; they require a replay verifier or remain pending.',
                            'Detected warnings may require a focused correction to the auditor if they are false positives; do not erase findings or substitute a new source baseline.']}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--source', type=Path, required=True)
    parser.add_argument('--final', type=Path, required=True)
    parser.add_argument('--report', type=Path, required=True)
    for name in ('chain', 'render-manifest', 'page-manifest', 'review', 'coverage'):
        parser.add_argument('--' + name, type=Path)
    args = parser.parse_args()
    try:
        inputs = [p.resolve() for p in (args.source, args.final, args.chain, args.render_manifest, args.page_manifest, args.review, args.coverage) if p]
        report = args.report.expanduser().resolve()
        if report in inputs or report.exists() or report.is_symlink():
            raise ValueError('Report must be a new path distinct from every input')
        if args.source.suffix.lower() != '.docx' or args.final.suffix.lower() != '.docx':
            raise ValueError('Source and final must be DOCX files')
        result = check(args.source.resolve(), args.final.resolve(),
                       **{key: getattr(args, key).resolve() if getattr(args, key) else None
                          for key in ('chain', 'render_manifest', 'page_manifest', 'review', 'coverage')})
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
