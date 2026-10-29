#!/usr/bin/env python3
"""Inspect review markup; derive a comment-free copy of the known school template.

The cleaning command is deliberately restricted to the retained source template.
Manuscript comments may be authored by a student or supervisor and are not removed
by this tool. Tracked revisions are counted by exact QName and never accepted.
"""
from __future__ import annotations

import argparse
from collections import Counter
import hashlib
import json
from pathlib import Path, PurePosixPath
import posixpath
import sys
from urllib.parse import unquote, urlsplit
import zipfile

from lxml import etree as E

TEMPLATE_SHA256 = '83dac94cb138ab7b46fb100d9ac05c6e1ab8c7dec986ebefd2e5ca22bf359eac'
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WORD_NAMESPACES = (W, 'http://purl.oclc.org/ooxml/wordprocessingml/main')
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
ANCHOR_NAMES = ('commentRangeStart', 'commentRangeEnd', 'commentReference')
REVISION_NAMES = ('ins', 'del', 'moveFrom', 'moveTo', 'pPrChange', 'rPrChange',
                  'sectPrChange', 'tblPrChange', 'tblGridChange', 'trPrChange',
                  'tcPrChange', 'numberingChange', 'cellIns', 'cellDel', 'cellMerge',
                  'moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart',
                  'moveToRangeEnd')
COMMENT_TYPES = {
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
    'http://purl.oclc.org/ooxml/officeDocument/relationships/comments',
    'http://schemas.microsoft.com/office/2011/relationships/commentsExtended',
    'http://schemas.microsoft.com/office/2016/09/relationships/commentsIds',
    'http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible',
    'http://schemas.microsoft.com/office/2011/relationships/people',
}
COMMENT_CONTENT_TYPES = {
    'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
    'application/vnd.ms-word.commentsExtended+xml',
    'application/vnd.ms-word.commentsIds+xml',
    'application/vnd.ms-word.commentsExtensible+xml',
    'application/vnd.ms-word.people+xml',
}
KNOWN_PART_NAMES = {'comments.xml', 'commentsExtended.xml', 'commentsIds.xml',
                    'commentsExtensible.xml', 'people.xml'}


def parse(data):
    return E.fromstring(data, E.XMLParser(resolve_entities=False, no_network=True))


def serialize(root):
    return E.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True)


def digest(data):
    return hashlib.sha256(data).hexdigest()


def read_package(path):
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        if len(names) != len(set(names)):
            raise ValueError('Duplicate ZIP members are not supported')
        return {name: archive.read(name) for name in names}


def normalized(text):
    return ''.join(text.split())


def template_comment_texts():
    source = Path(__file__).resolve().parents[1] / 'references' / 'source-comments.json'
    records = json.loads(source.read_text(encoding='utf-8-sig'))
    return {normalized(item['text']) for item in records if item.get('text')}


def word_name(node):
    if not isinstance(node.tag, str):
        return None
    name = E.QName(node)
    return name.localname if name.namespace in WORD_NAMESPACES else None


def word_attr(node, key):
    for namespace in WORD_NAMESPACES:
        value = node.get('{' + namespace + '}' + key)
        if value is not None:
            return value
    return None


def relationship_owner(name):
    if name == '_rels/.rels':
        return ''
    path = PurePosixPath(name)
    if path.parent.name != '_rels' or not path.name.endswith('.rels'):
        raise ValueError('Unexpected relationship part path: ' + name)
    return (path.parent.parent / path.name[:-5]).as_posix()


def relationship_target(rels_name, relation):
    if relation.get('TargetMode') == 'External':
        return None
    raw = relation.get('Target', '')
    if not raw:
        return None
    target = unquote(urlsplit(raw).path)
    owner = relationship_owner(rels_name)
    return posixpath.normpath(target.lstrip('/') if target.startswith('/') else
                             posixpath.join(posixpath.dirname(owner), target))


def own_relationship_part(part):
    path = PurePosixPath(part)
    return (path.parent / '_rels' / (path.name + '.rels')).as_posix()


def is_known_comment_part(name):
    return name.startswith('word/') and PurePosixPath(name).name in KNOWN_PART_NAMES


def comment_content_types(parts, discovered):
    """Include stale overrides even when both the part and its relation are gone."""
    records = []
    if '[Content_Types].xml' not in parts:
        return records
    for item in parse(parts['[Content_Types].xml']):
        if item.tag != '{' + CT + '}Override':
            continue
        target = unquote(item.get('PartName', '')).lstrip('/')
        content_type = item.get('ContentType')
        if content_type in COMMENT_CONTENT_TYPES or target in discovered or is_known_comment_part(target):
            records.append({'part': target, 'content_type': content_type,
                            'missing_target': target not in parts})
    return records


def comment_parts(parts):
    """Discover standard comment metadata by name, type and relationships."""
    found = {name for name in parts if is_known_comment_part(name)}
    if '[Content_Types].xml' in parts:
        for item in parse(parts['[Content_Types].xml']):
            if item.tag == '{' + CT + '}Override' and item.get('ContentType') in COMMENT_CONTENT_TYPES:
                target = unquote(item.get('PartName', '')).lstrip('/')
                if target in parts:
                    found.add(target)
    for name, data in parts.items():
        if not name.endswith('.rels'):
            continue
        for relation in parse(data):
            if relation.tag == '{' + PKG + '}Relationship' and relation.get('Type') in COMMENT_TYPES:
                target = relationship_target(name, relation)
                if target in parts:
                    found.add(target)
    return found


def inspect_parts(parts, template_texts=None):
    """Read-only metadata for Audit; no content changes or visual claims.

    template_comment_matches is evidence of identical historical wording, not
    permission to delete arbitrary manuscript comments. Counts describe XML
    nodes, not distinct human editing operations.
    """
    known = template_comment_texts() if template_texts is None else {normalized(t) for t in template_texts}
    counts = Counter({name: 0 for name in ANCHOR_NAMES})
    revisions = Counter({name: 0 for name in REVISION_NAMES})
    comments = []
    anchor_ids = set()
    missing_id_anchors = 0
    matches = []
    discovered = comment_parts(parts)
    content_types = comment_content_types(parts, discovered)
    relationships = []
    for part, data in parts.items():
        if not part.endswith(('.xml', '.rels')):
            continue
        root = parse(data)
        for node in root.iter():
            name = word_name(node)
            if name == 'comment':
                value = ''.join(n.text or '' for n in node.iter() if word_name(n) == 't')
                record = {'part': part, 'id': word_attr(node, 'id'), 'text': value}
                comments.append(record)
                if normalized(value) and normalized(value) in known:
                    matches.append(record)
            elif name in ANCHOR_NAMES:
                counts[name] += 1
                ident = word_attr(node, 'id')
                if ident is None:
                    missing_id_anchors += 1
                else:
                    anchor_ids.add(ident)
            elif name in REVISION_NAMES:
                revisions[name] += 1
        if part.endswith('.rels'):
            for relation in root:
                if relation.tag != '{' + PKG + '}Relationship':
                    continue
                target = relationship_target(part, relation)
                if relation.get('Type') in COMMENT_TYPES or target in discovered:
                    relationships.append({'part': part, 'id': relation.get('Id'),
                                          'type': relation.get('Type'), 'target': relation.get('Target'),
                                          'missing_target': target is not None and target not in parts})
    comment_ids = {item['id'] for item in comments if item['id'] is not None}
    return {'comment_count': len(comments), 'template_comment_matches': matches,
            'anchors': dict(counts), 'comment_parts': sorted(discovered),
            'comment_relationships': relationships, 'revisions': dict(revisions),
            'comment_content_types': content_types,
            'dangling_comment_content_types': [item for item in content_types if item['missing_target']],
            'dangling_comment_relationships': [item for item in relationships if item['missing_target']],
            'orphan_anchor_ids': sorted(anchor_ids - comment_ids),
            'missing_id_anchors': missing_id_anchors}


def strip_comment_parts(parts):
    """Internal deterministic transform. Remove comments only, never revisions.

    Caller is responsible for authorization/provenance. The public CLI allows
    this transform only on the SHA-256-pinned school template.
    """
    removed = comment_parts(parts)
    content_type_targets = {item['part'] for item in comment_content_types(parts, removed)}
    removed |= {own_relationship_part(name) for name in tuple(removed)
                if own_relationship_part(name) in parts}
    result = {name: data for name, data in parts.items() if name not in removed}
    changes = {'removed_parts': sorted(removed), 'removed_anchors': 0,
               'removed_relationships': 0, 'removed_content_types': 0, 'changed_parts': []}
    for part, data in tuple(result.items()):
        if not part.endswith(('.xml', '.rels')):
            continue
        root = parse(data)
        changed = False
        for node in tuple(root.iter()):
            if word_name(node) in ANCHOR_NAMES:
                parent = node.getparent()
                if parent is None:
                    raise ValueError('Comment anchor cannot be the document root')
                # Do not remove the enclosing run: it may contain a field/image.
                parent.remove(node)
                changes['removed_anchors'] += 1
                changed = True
        if part.endswith('.rels'):
            for relation in tuple(root):
                if relation.tag == '{' + PKG + '}Relationship' and (
                        relation.get('Type') in COMMENT_TYPES or
                        relationship_target(part, relation) in removed):
                    root.remove(relation)
                    changes['removed_relationships'] += 1
                    changed = True
        elif part == '[Content_Types].xml':
            for item in tuple(root):
                if item.tag == '{' + CT + '}Override' and (
                        unquote(item.get('PartName', '')).lstrip('/') in removed | content_type_targets or
                        item.get('ContentType') in COMMENT_CONTENT_TYPES):
                    root.remove(item)
                    changes['removed_content_types'] += 1
                    changed = True
        if changed:
            result[part] = serialize(root)
            changes['changed_parts'].append(part)
    return result, changes


def inspect(path):
    path = Path(path)
    result = inspect_parts(read_package(path))
    result['source_sha256'] = digest(path.read_bytes())
    return result


def clean_template(source, output, report=None):
    source, output = Path(source).resolve(), Path(output).resolve()
    report = Path(report).resolve() if report else None
    if source == output or output.exists():
        raise ValueError('Output must be a new path distinct from the source')
    if report and (report in {source, output} or report.exists()):
        raise ValueError('Report must be a new path distinct from source and output')
    original = source.read_bytes()
    if digest(original) != TEMPLATE_SHA256:
        raise ValueError('clean-template accepts only the pinned original school template; manuscripts are read-only')
    parts = read_package(source)
    cleaned, changes = strip_comment_parts(parts)
    before, after = inspect_parts(parts), inspect_parts(cleaned)
    if (after['comment_count'] or after['comment_parts'] or after['comment_relationships'] or
            after['comment_content_types'] or any(after['anchors'].values())):
        raise ValueError('Comment cleanup was incomplete')
    if before['revisions'] != after['revisions']:
        raise ValueError('Tracked revision nodes changed; cleanup refused')
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, 'x', zipfile.ZIP_DEFLATED) as archive:
        for name in sorted(cleaned):
            # Stable ZIP metadata makes the derived asset reproducible.
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.create_system = 3
            info.external_attr = 0o600 << 16
            archive.writestr(info, cleaned[name])
    if source.read_bytes() != original:
        raise ValueError('Source changed during cleanup')
    result = {'operation_succeeded': True, 'source_sha256': TEMPLATE_SHA256,
              'output_sha256': digest(output.read_bytes()), 'changes': changes,
              'before': before, 'after': after,
              'scope': 'template_comments_only', 'visual_review_performed': False}
    if report:
        report.parent.mkdir(parents=True, exist_ok=True)
        with report.open('x', encoding='utf-8') as handle:
            json.dump(result, handle, ensure_ascii=False, indent=2)
            handle.write('\n')
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest='command', required=True)
    check = commands.add_parser('inspect', help='Read-only exact XML review-markup inventory')
    check.add_argument('source', type=Path)
    clean = commands.add_parser('clean-template', help='Derive a clean copy of the pinned source template')
    clean.add_argument('source', type=Path)
    clean.add_argument('output', type=Path)
    clean.add_argument('--report', type=Path)
    args = parser.parse_args()
    try:
        result = inspect(args.source) if args.command == 'inspect' else clean_template(args.source, args.output, args.report)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0
    except (ValueError, OSError, zipfile.BadZipFile, E.XMLSyntaxError) as exc:
        print(json.dumps({'operation_succeeded': False, 'error': str(exc)}, ensure_ascii=False))
        return 2


if __name__ == '__main__':
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    raise SystemExit(main())
