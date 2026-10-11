#!/usr/bin/env python3
"""Create bounded, source-bound DOCX review batches and verify complete coverage.

Coverage records decisions, not the quality of a human/agent's reasoning or the
visual correctness of the finished thesis. This tool never edits a DOCX.
"""
from __future__ import annotations

import argparse
import copy
import json
from pathlib import Path
import sys

import thesis_format as F

SCHEMA_VERSION = 1
MAIN = 'word/document.xml'
KINDS = ('paragraph', 'table', 'story_paragraph', 'story_table')


def _json_bytes(value):
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(',', ':')).encode('utf-8')


def _sha_file(path):
    try:
        return F.sha(Path(path).read_bytes())
    except OSError as exc:
        raise ValueError('Cannot hash required review evidence: ' + str(path)) from exc


def _load(path):
    # Duplicate JSON keys must not silently replace a reviewed decision.
    def unique(pairs):
        value = {}
        for key, item in pairs:
            if key in value:
                raise ValueError('Duplicate JSON key: ' + key)
            value[key] = item
        return value
    try:
        return json.loads(Path(path).read_text(encoding='utf-8-sig'), object_pairs_hook=unique)
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError('Cannot read review evidence: ' + str(path)) from exc


def _save_new(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('x', encoding='utf-8') as handle:
        json.dump(value, handle, ensure_ascii=False, indent=2)
        handle.write('\n')


def _require(condition, message):
    if not condition:
        raise ValueError(message)


def _inside(folder, relative):
    _require(isinstance(relative, str), 'Evidence path must be a string')
    path = (folder / relative).resolve()
    _require(path.is_relative_to(folder.resolve()), 'Evidence path leaves the review directory')
    return path


def _source_items(source):
    """Use the same per-part IDs and exact ptext as thesis_format.inspect_doc."""
    source = Path(source).expanduser().resolve()
    before = _sha_file(source)
    parts = F.package(source)
    _require(MAIN in parts, 'Source has no word/document.xml')
    items = []
    for part in [MAIN] + sorted(n for n in parts if n != MAIN and n.startswith('word/') and n.endswith('.xml')):
        root = F.parse(parts[part])
        ps, ts = F.indexed(root)
        if not ps and not ts:
            continue
        ids = {node: pid for pid, node in ps.items()} | {node: tid for tid, node in ts.items()}
        for node in root.iter():
            if node not in ids:
                continue
            is_table = node.tag == F.q('tbl')
            protected = any(a.tag in F.PROTECTED_TAGS for a in node.iterancestors())
            hints = {
                'inside_protected_object': protected,
                'has_drawing': any(e.tag in {F.q('drawing'), F.q('pict')} for e in node.iter()),
                'has_math': any(e.tag == '{' + F.M + '}oMath' for e in node.iter()),
                'has_object': any(e.tag == F.q('object') for e in node.iter()),
                'has_field': any(e.tag in {F.q('fldChar'), F.q('fldSimple')} for e in node.iter()),
                'in_textbox': any(F.E.QName(a).localname == 'txbxContent' for a in node.iterancestors()),
            }
            ancestor_table = next((a for a in node.iterancestors() if a.tag == F.q('tbl')), None)
            st = node.find('w:pPr/w:pStyle', F.NS)
            text = F.ptext(node)
            item = {'key': part + '#' + ids[node], 'part': part, 'id': ids[node],
                    'ordinal': len(items) + 1,
                    'kind': ('story_' if part != MAIN else '') + ('table' if is_table else 'paragraph'),
                    'text': text, 'text_sha256': F.sha(text.encode('utf-8')),
                    'content_sha256': F.sha(_json_bytes(F.tree_signature(node))),
                    'style': st.get(F.q('val')) if st is not None else None,
                    'table_id': ids.get(ancestor_table), 'object_hints': hints,
                    'format_supported': not protected and (part == MAIN or (not is_table and F.allowed_story(part)))}
            if is_table:
                item['rows'] = len(node.findall('w:tr', F.NS))
                item['paragraph_ids'] = [pid for pid, p in ps.items()
                                         if next((a for a in p.iterancestors() if a.tag == F.q('tbl')), None) is node]
            items.append(item)
    _require(_sha_file(source) == before, 'Source changed while building its review index')
    return before, items, parts


def _light(item):
    return {k: item[k] for k in ('key', 'part', 'id', 'ordinal', 'kind', 'text_sha256', 'content_sha256')} | {'text_characters': len(item['text'])}


def _context(item):
    if item is None:
        return None
    return {'key': item['key'], 'text': item['text'][:500], 'text_sha256': item['text_sha256'],
            'truncated': len(item['text']) > 500, 'context_only': True}


def _groups(items, chunk_size, max_chars):
    groups, group, chars = [], [], 0
    for item in items:
        length = len(item['text'])
        if group and (len(group) >= chunk_size or chars + length > max_chars):
            groups.append(group)
            group, chars = [], 0
        group.append(item)
        chars += length
    if group:
        groups.append(group)
    return groups


def create_review_plan(source, work_dir, chunk_size=120, draft_path=None, max_chars=60000):
    """Create a new directory; source and an optional existing draft stay intact."""
    _require(type(chunk_size) is int and 1 <= chunk_size <= 150, 'chunk_size must be 1..150 (default 120)')
    _require(type(max_chars) is int and max_chars > 0, 'max_chars must be a positive integer')
    folder = Path(work_dir).expanduser().resolve()
    _require(not folder.exists(), 'Review directory must be new; no outputs will be overwritten')
    source = Path(source).expanduser().resolve()
    source_sha, items, _ = _source_items(source)
    _require(items, 'Source contains no reviewable paragraph or table')
    if draft_path is None:
        draft = {'source_sha256': source_sha, 'reviewed': False, 'paragraphs': {}, 'tables': {}, 'story_paragraphs': {}}
        draft_bytes = _json_bytes(draft)
        draft_origin = None
    else:
        draft_path = Path(draft_path).expanduser().resolve()
        draft_bytes = draft_path.read_bytes()
        draft = _load(draft_path)
        _require(draft.get('source_sha256') == source_sha, 'Draft source_sha256 is missing or stale')
        draft_origin = str(draft_path)
    folder.mkdir(parents=True, exist_ok=False)
    with (folder / 'style-map.draft.json').open('xb') as handle:
        handle.write(draft_bytes)
    groups = _groups(items, chunk_size, max_chars)
    index = {'schema_version': SCHEMA_VERSION, 'source': str(source), 'source_sha256': source_sha,
             'chunk_size': chunk_size, 'max_chars': max_chars,
             'draft_source': {'original_path': draft_origin, 'path': 'style-map.draft.json', 'sha256': F.sha(draft_bytes)},
             'items': [_light(item) for item in items], 'chunks': []}
    offset = 0
    for number, group in enumerate(groups, 1):
        name = f'chunks/chunk-{number:04d}.json'
        context = {'before': _context(items[offset - 1] if offset else None),
                   'after': _context(items[offset + len(group)] if offset + len(group) < len(items) else None)}
        chunk = {'schema_version': SCHEMA_VERSION, 'source_sha256': source_sha, 'chunk_id': f'chunk-{number:04d}',
                 'context': context, 'items': [item | {'decision': {'reviewed': False, 'action': None, 'role': None, 'reason': ''}} for item in group]}
        _save_new(folder / name, chunk)
        index['chunks'].append({'chunk_id': chunk['chunk_id'], 'path': name,
                                'keys': [item['key'] for item in group],
                                'source_content_sha256': F.sha(_json_bytes({'context': context, 'items': group}))})
        offset += len(group)
    _save_new(folder / 'index.json', index)
    return {'operation_succeeded': True, 'delivery_ready': False, 'status': 'awaiting_full_review',
            'work_directory': str(folder), 'index_path': str(folder / 'index.json'),
            'source_sha256': source_sha, 'counts': _counts(items, [], len(groups)),
            'next_chunk': str(folder / index['chunks'][0]['path'])}


def _counts(items, decisions, chunks):
    reviewed = sum(d.get('reviewed') is True and d.get('action') in ('format', 'preserve') for d in decisions)
    return {'total': len(items), **{kind + 's': sum(i['kind'] == kind for i in items) for kind in KINDS},
            'reviewed': reviewed, 'unreviewed': len(items) - reviewed,
            'formatted': sum(d.get('action') == 'format' for d in decisions),
            'preserved': sum(d.get('action') == 'preserve' for d in decisions), 'chunks': chunks}


def _decision(item, decision):
    label = item['key']
    _require(isinstance(decision, dict) and not set(decision) - {'reviewed', 'action', 'role', 'reason', 'header_rows'}, 'Invalid decision fields: ' + label)
    _require(decision.get('reviewed') is True, 'Unreviewed item: ' + label)
    action = decision.get('action')
    _require(action in ('format', 'preserve'), 'Explicit format or preserve decision required: ' + label)
    if action == 'preserve':
        reason = decision.get('reason')
        _require(isinstance(reason, str) and len(reason.strip()) >= 6 and reason.strip().lower() not in {'preserve', 'unchanged', 'no change', '保持原样', '无需修改'},
                 'Preserve requires a specific reason (at least 6 characters): ' + label)
        _require(decision.get('role') in (None, '') and 'header_rows' not in decision, 'Preserve cannot also assign formatting: ' + label)
    else:
        _require(item['format_supported'], 'Protected or unsupported story item must be preserved with a reason: ' + label)
        role = decision.get('role')
        if item['kind'] in ('table', 'story_table'):
            _require(role in ('three-line', 'complex'), 'Table role must be three-line or complex: ' + label)
            rows = decision.get('header_rows', 1)
            _require(type(rows) is int and 0 <= rows <= item['rows'], 'Invalid table header_rows: ' + label)
        else:
            _require(isinstance(role, str) and role in F.ROLES, 'Unknown paragraph role: ' + label)
            _require('header_rows' not in decision, 'header_rows is only for tables: ' + label)
    return copy.deepcopy(decision)


def _read_plan(source, folder, require_reviewed=True):
    folder = Path(folder).expanduser().resolve()
    index = _load(folder / 'index.json')
    source_sha, items, parts = _source_items(source)
    _require(index.get('schema_version') == SCHEMA_VERSION, 'Unsupported review index schema_version')
    _require(index.get('source_sha256') == source_sha, 'Source SHA changed; create a new review plan for the exact source')
    _require(index.get('items') == [_light(item) for item in items], 'Index does not match the full source item set/order')
    draft = index.get('draft_source', {})
    draft_path = _inside(folder, draft.get('path'))
    _require(_sha_file(draft_path) == draft.get('sha256'), 'Retained draft changed')
    _require(_load(draft_path).get('source_sha256') == source_sha, 'Retained draft is bound to a different source')
    entries = index.get('chunks')
    _require(isinstance(entries, list) and entries, 'Review index has no chunks')
    _require(type(index.get('chunk_size')) is int and 1 <= index['chunk_size'] <= 150 and
             type(index.get('max_chars')) is int and index['max_chars'] > 0, 'Invalid review chunk limits')
    expected_groups = _groups(items, index.get('chunk_size', 0), index.get('max_chars', 0))
    _require(len(entries) == len(expected_groups), 'Missing or extra chunk in the review index')
    records, evidence, seen, pending = [], [], set(), []
    offset = 0
    for number, (entry, group) in enumerate(zip(entries, expected_groups), 1):
        chunk_id = f'chunk-{number:04d}'
        _require(entry.get('chunk_id') == chunk_id and entry.get('path') == f'chunks/{chunk_id}.json', 'Chunk identity/order changed')
        _require(entry.get('keys') == [i['key'] for i in group], 'Chunk partition dropped, duplicated or reordered source items')
        path = _inside(folder, entry['path'])
        chunk = _load(path)
        _require(chunk.get('schema_version') == SCHEMA_VERSION and chunk.get('source_sha256') == source_sha and chunk.get('chunk_id') == chunk_id, 'Chunk source/identity mismatch: ' + chunk_id)
        context = {'before': _context(items[offset - 1] if offset else None),
                   'after': _context(items[offset + len(group)] if offset + len(group) < len(items) else None)}
        _require(chunk.get('context') == context, 'Chunk boundary context changed: ' + chunk_id)
        actual = chunk.get('items')
        _require(isinstance(actual, list) and len(actual) == len(group), 'Missing or extra item in ' + chunk_id)
        for record, expected in zip(actual, group):
            _require(isinstance(record, dict), 'Invalid review item in ' + chunk_id)
            key = record.get('key')
            _require(isinstance(key, str) and key not in seen, 'Duplicate item: ' + str(key))
            seen.add(key)
            _require({k: v for k, v in record.items() if k != 'decision'} == expected, 'Unknown, changed or reordered source item: ' + str(key))
            decision = record.get('decision')
            if require_reviewed:
                decision = _decision(expected, decision)
            elif not isinstance(decision, dict) or decision.get('reviewed') is not True or decision.get('action') not in ('format', 'preserve'):
                if path not in pending:
                    pending.append(path)
            else:
                _decision(expected, decision)
            records.append({'key': key, 'text_sha256': expected['text_sha256'], 'content_sha256': expected['content_sha256'], 'decision': decision})
        content_sha = F.sha(_json_bytes({'context': context, 'items': group}))
        _require(entry.get('source_content_sha256') == content_sha, 'Chunk source-content hash mismatch: ' + chunk_id)
        evidence.append({'chunk_id': chunk_id, 'path': str(path), 'sha256': _sha_file(path), 'items': len(group)})
        offset += len(group)
    _require(len(seen) == len(items), 'Review did not cover every source item')
    return index, items, parts, records, evidence, pending


def _mapping(source, source_sha, items, parts, records):
    mapping = {'schema_version': SCHEMA_VERSION, 'source_sha256': source_sha, 'reviewed': True,
               'page_setup': False, 'paragraphs': {}, 'tables': {}, 'story_paragraphs': {}}
    for item, record in zip(items, records):
        decision = record['decision']
        if decision['action'] == 'preserve':
            continue
        role = decision['role']
        if item['kind'] == 'table':
            mapping['tables'][item['id']] = {'kind': role, 'header_rows': decision.get('header_rows', 1)}
        else:
            target = mapping['paragraphs'] if item['part'] == MAIN else mapping['story_paragraphs'].setdefault(item['part'], {})
            target[item['id']] = {'role': role, 'expected_text': item['text']}
    ps, ts = F.indexed(F.parse(parts[MAIN]))
    if mapping['paragraphs'] or mapping['tables'] or mapping['story_paragraphs']:
        F.validate_map(mapping, source, ps, ts, parts)
    return mapping


def assemble_review_plan(source, work_dir, style_map_path, coverage_path):
    """Fail closed until every source item has a valid explicit reviewed decision."""
    folder = Path(work_dir).expanduser().resolve()
    outputs = [Path(p).expanduser().resolve() for p in (style_map_path, coverage_path)]
    _require(outputs[0] != outputs[1] and all(not p.exists() for p in outputs), 'Outputs must be distinct new paths; no files will be overwritten')
    index, items, parts, records, evidence, _ = _read_plan(source, folder)
    provenance = {'index_path': str(folder / 'index.json'), 'index_sha256': _sha_file(folder / 'index.json'),
                  'draft_source': index['draft_source']}
    mapping = _mapping(source, index['source_sha256'], items, parts, records)
    mapping['review_provenance'] = provenance
    _save_new(outputs[0], mapping)
    coverage = {'schema_version': SCHEMA_VERSION, 'source_sha256': index['source_sha256'],
                'reviewed': True, 'coverage_complete': True, 'delivery_ready': False,
                'review_directory': str(folder), **provenance,
                'style_map_path': str(outputs[0]), 'style_map_sha256': _sha_file(outputs[0]),
                'chunks': evidence, 'counts': _counts(items, [r['decision'] for r in records], len(evidence)),
                'decisions': records,
                'limitations': ['Coverage verifies recorded decisions, not reviewer comprehension or role correctness.',
                                'Coverage does not establish layout correctness or visual acceptance.',
                                'Preserve decisions do not waive content preservation checks.']}
    _save_new(outputs[1], coverage)
    return {'operation_succeeded': True, 'delivery_ready': False, 'status': 'review_assembled',
            'style_map_path': str(outputs[0]), 'coverage_path': str(outputs[1]), 'counts': coverage['counts'],
            'mapped_item_count': coverage['counts']['formatted'], 'preserve_count': coverage['counts']['preserved'],
            'apply_required': bool(coverage['counts']['formatted']),
            'next_step': 'Apply this reviewed map to the exact source.' if coverage['counts']['formatted'] else
                         'All items are explicitly preserved. Do not apply the empty map; formatting acceptance still requires audit and visual review.'}


def verify_coverage(source, coverage_path):
    """Recompute all evidence against the ORIGINAL source; raise ValueError on failure."""
    coverage = _load(coverage_path)
    _require(coverage.get('schema_version') == SCHEMA_VERSION and coverage.get('reviewed') is True and coverage.get('coverage_complete') is True, 'Coverage must have a supported schema and complete reviewed status')
    folder = Path(coverage.get('review_directory', '')).expanduser().resolve()
    index, items, parts, records, evidence, _ = _read_plan(source, folder)
    _require(coverage.get('source_sha256') == index['source_sha256'], 'Coverage source_sha256 mismatch')
    _require(coverage.get('index_path') == str(folder / 'index.json') and coverage.get('index_sha256') == _sha_file(folder / 'index.json'), 'Coverage index evidence changed')
    _require(coverage.get('draft_source') == index['draft_source'], 'Coverage draft provenance changed')
    _require(coverage.get('chunks') == evidence, 'Coverage chunk hashes/evidence changed')
    _require(coverage.get('decisions') == records, 'Coverage decisions do not match the full reviewed source item set')
    counts = _counts(items, [r['decision'] for r in records], len(evidence))
    _require(coverage.get('counts') == counts, 'Coverage counts mismatch')
    style_map_path = Path(coverage.get('style_map_path', '')).expanduser().resolve()
    _require(coverage.get('style_map_sha256') == _sha_file(style_map_path), 'Assembled style map hash changed')
    expected = _mapping(source, index['source_sha256'], items, parts, records)
    expected['review_provenance'] = {'index_path': str(folder / 'index.json'), 'index_sha256': coverage['index_sha256'], 'draft_source': index['draft_source']}
    _require(_load(style_map_path) == expected, 'Assembled map does not match reviewed decisions')
    return {'coverage_complete': True, 'source_sha256': index['source_sha256'], 'counts': counts,
            'mapped_item_count': counts['formatted'], 'preserve_count': counts['preserved'],
            'style_map_path': str(style_map_path), 'style_map_sha256': coverage['style_map_sha256'],
            'verifiable': {'source_sha256': True, 'all_source_items': True, 'unique_ids': True,
                           'original_text_and_objects': True, 'all_decisions_reviewed': True,
                           'chunk_hashes': True, 'style_map_matches_decisions': True, 'draft_retained': True},
            'delivery_ready': False}


def next_chunk(source, work_dir):
    """Return only one outstanding batch; repeat after saving its reviewed decisions."""
    _, items, _, records, evidence, pending = _read_plan(source, work_dir, require_reviewed=False)
    return {'coverage_complete': not pending, 'counts': _counts(items, [r['decision'] or {} for r in records], len(evidence)),
            'pending_chunks': len(pending), 'chunk_path': str(pending[0]) if pending else None,
            'chunk': _load(pending[0]) if pending else None}


def review_status(source, work_dir):
    """Return a compact continuation checkpoint without including paragraph text."""
    _, items, _, records, evidence, pending = _read_plan(source, work_dir, require_reviewed=False)
    return {'coverage_complete': not pending, 'delivery_ready': False,
            'counts': _counts(items, [r['decision'] or {} for r in records], len(evidence)),
            'pending_chunks': [str(path) for path in pending],
            'next_chunk': str(pending[0]) if pending else None}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='command', required=True)
    for name in ('create', 'status', 'next', 'assemble', 'verify'):
        command = sub.add_parser(name)
        command.add_argument('--source', type=Path, required=True)
        if name != 'verify':
            command.add_argument('--work-dir', type=Path, required=True)
        if name == 'create':
            command.add_argument('--chunk-size', type=int, default=120)
            command.add_argument('--max-chars', type=int, default=60000)
            command.add_argument('--draft', type=Path)
        if name in ('assemble', 'verify'):
            command.add_argument('--coverage', type=Path, required=True)
        if name == 'assemble':
            command.add_argument('--style-map', type=Path, required=True)
    args = parser.parse_args()
    try:
        if args.command == 'create':
            result = create_review_plan(args.source, args.work_dir, args.chunk_size, args.draft, args.max_chars)
        elif args.command == 'next':
            result = next_chunk(args.source, args.work_dir)
        elif args.command == 'status':
            result = review_status(args.source, args.work_dir)
        elif args.command == 'assemble':
            result = assemble_review_plan(args.source, args.work_dir, args.style_map, args.coverage)
        else:
            result = verify_coverage(args.source, args.coverage)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0
    except (ValueError, OSError, TypeError, KeyError) as exc:
        print(json.dumps({'operation_succeeded': False, 'delivery_ready': False, 'error': str(exc)}, ensure_ascii=False))
        return 2


if __name__ == '__main__':
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    raise SystemExit(main())
