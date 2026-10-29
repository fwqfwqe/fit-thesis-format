#!/usr/bin/env python3
"""Get the complete skill with Python's standard library; do not format papers."""
import argparse
import hashlib
import json
from pathlib import Path, PurePosixPath
import sys
import urllib.request
import zipfile
import io

RAW = 'https://raw.githubusercontent.com/fwqfwqe/fit-thesis-format/main/'


def sha(data):
    return hashlib.sha256(data).hexdigest()


def fetch(relative):
    request = urllib.request.Request(RAW + relative, headers={'User-Agent': 'fit-thesis-format-bootstrap'})
    with urllib.request.urlopen(request, timeout=45) as response:
        data = response.read(25 * 1024 * 1024 + 1)
    if len(data) > 25 * 1024 * 1024:
        raise ValueError('Download exceeds the package size limit')
    return data


def safe_name(value):
    if not isinstance(value, str) or '\\' in value or ':' in value:
        raise ValueError('Invalid package path')
    path = PurePosixPath(value)
    if path.is_absolute() or '..' in path.parts or not path.parts:
        raise ValueError('Invalid package path')
    return path


def install(dest, manifest_bytes, archive_bytes):
    if dest.exists() or dest.is_symlink():
        raise ValueError('Destination must be a new directory; nothing will be overwritten')
    manifest = json.loads(manifest_bytes.decode('utf-8-sig'))
    if manifest.get('skill') != 'fit-thesis-format':
        raise ValueError('Unexpected skill name')
    archive = manifest['archive']
    if len(archive_bytes) != archive['size'] or sha(archive_bytes) != archive['sha256']:
        raise ValueError('Archive size/SHA-256 does not match the manifest')
    entries = manifest['files']
    expected = {entry['path']: entry for entry in entries}
    if len(expected) != len(entries):
        raise ValueError('Duplicate manifest entries')
    payloads = []
    with zipfile.ZipFile(io.BytesIO(archive_bytes)) as bundle:
        if len(bundle.namelist()) != len(set(bundle.namelist())) or set(bundle.namelist()) != set(expected):
            raise ValueError('ZIP members do not exactly match the manifest')
        if sum(i.file_size for i in bundle.infolist()) > 80 * 1024 * 1024:
            raise ValueError('Expanded package too large')
        for item in bundle.infolist():
            path = safe_name(item.filename)
            if path.parts[0] != 'fit-thesis-format' or len(path.parts) < 2:
                raise ValueError('Unexpected package root')
            if item.is_dir() or (item.external_attr >> 16) & 0o170000 == 0o120000:
                raise ValueError('Directories/symlinks must not be package members')
            data = bundle.read(item.filename)
            entry = expected[item.filename]
            if len(data) != entry['size'] or sha(data) != entry['sha256']:
                raise ValueError('Package member failed verification: ' + item.filename)
            payloads.append((path, data))
    for required in ('SKILL.md', 'START_HERE.md', 'scripts/thesis_job.py', 'scripts/format_engine.py',
                     'scripts/integrity.py', 'scripts/office_render.ps1', 'references/rules-v10.json',
                     'scripts/prepare.py', 'scripts/delivery_check.py', 'scripts/review_plan.py',
                     'scripts/narrative_citations.py', 'references/citation-consistency.md',
                     'scripts/template_hygiene.py', 'references/template-hygiene.md',
                     'assets/fit-economics-management-template-clean.docx',
                     'assets/fit-economics-management-template.docx', 'references/style-spec.json'):
        if 'fit-thesis-format/' + required not in expected:
            raise ValueError('Incomplete current package: ' + required)
    dest.mkdir(parents=True, exist_ok=False)
    for path, data in payloads:
        target = dest.joinpath(*path.parts)
        target.parent.mkdir(parents=True, exist_ok=True)
        with target.open('xb') as handle:
            handle.write(data)
    with (dest / 'download-manifest.json').open('xb') as handle:
        handle.write(manifest_bytes)
    return {'operation_succeeded': True, 'status': 'package_downloaded_and_verified',
            'version': manifest['version'], 'file_count': len(entries),
            'skill_directory': str(dest / 'fit-thesis-format'),
            'next_read': str(dest / 'fit-thesis-format' / 'START_HERE.md'),
            'delivery_ready': False,
            'remaining_steps': ['Read START_HERE.md and SKILL.md', 'Run thesis_job.py with the real paper',
                                'Complete semantic and visual tasks, then resume',
                                'Verify the same final paper in Word and WPS']}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--dest', type=Path, required=True)
    parser.add_argument('--manifest', type=Path, help='Offline manifest; supply with --archive')
    parser.add_argument('--archive', type=Path, help='Offline ZIP; supply with --manifest')
    args = parser.parse_args()
    try:
        if bool(args.manifest) != bool(args.archive):
            raise ValueError('--manifest and --archive must be provided together')
        if args.dest.exists():
            raise ValueError('Destination already exists')
        manifest_bytes = args.manifest.read_bytes() if args.manifest else fetch('package-manifest.json')
        manifest = json.loads(manifest_bytes.decode('utf-8-sig'))
        archive_path = safe_name(manifest['archive']['path']).as_posix()
        if not archive_path.startswith('dist/') or not archive_path.endswith('.zip'):
            raise ValueError('Unexpected archive path')
        archive_bytes = args.archive.read_bytes() if args.archive else fetch(archive_path)
        result = install(args.dest.expanduser().resolve(), manifest_bytes, archive_bytes)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0
    except Exception as exc:
        print(json.dumps({'operation_succeeded': False, 'delivery_ready': False, 'error': str(exc)}, ensure_ascii=False))
        return 2


if __name__ == '__main__':
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    raise SystemExit(main())
