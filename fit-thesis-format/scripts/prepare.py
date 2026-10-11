#!/usr/bin/env python3
"""Prepare a real paper for review without modifying it or guessing role mappings."""
import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parent.parent


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def save(path, data):
    with Path(path).open('x', encoding='utf-8') as handle:
        json.dump(data, handle, ensure_ascii=False, indent=2)
        handle.write('\n')


def invoke(request):
    run = subprocess.run([sys.executable, str(ROOT / 'scripts' / 'agent_cli.py'), '--request', '-'],
                         input=json.dumps(request, ensure_ascii=False), text=True, encoding='utf-8', capture_output=True)
    try:
        result = json.loads(run.stdout)
    except ValueError as exc:
        raise ValueError('Tool did not return JSON: ' + run.stderr[-600:]) from exc
    if run.returncode or result.get('ok') is not True:
        raise ValueError('Preparation step failed: ' + json.dumps(result, ensure_ascii=False))
    return result


def prepare(source, work):
    source = source.expanduser().resolve()
    work = work.expanduser().resolve()
    if not source.is_file() or source.suffix.lower() != '.docx':
        raise ValueError('Source must be a readable DOCX file')
    if work.exists() or work.is_symlink():
        raise ValueError('Work directory must be new; no files will be overwritten')
    required = ['SKILL.md', 'START_HERE.md', 'agent-tools.json', 'requirements.txt',
                'references/style-spec.json', 'references/workflow.md',
                'references/format-requirements.md', 'scripts/review_plan.py',
                'assets/fit-economics-management-template.docx']
    missing = [name for name in required if not (ROOT / name).is_file()]
    if missing:
        raise ValueError('Incomplete skill package: ' + ', '.join(missing))
    original_sha = sha(source)
    template_sha = sha(ROOT / 'assets' / 'fit-economics-management-template.docx')
    if template_sha != '83dac94cb138ab7b46fb100d9ac05c6e1ab8c7dec986ebefd2e5ca22bf359eac':
        raise ValueError('School template does not match the retained reference')
    # Discover before creating outputs. The doctor does not claim a renderer works.
    environment = invoke({'action': 'doctor'})
    if environment.get('result', {}).get('static_cli_ready') is not True:
        raise ValueError('Static tooling is not ready. Use Python 3.10+ and install the package requirements before preparing a paper; no work directory was created.')
    work.mkdir(parents=True, exist_ok=False)
    save(work / 'doctor.json', environment)
    backup = work / 'original.docx'
    with backup.open('xb') as handle:
        handle.write(source.read_bytes())
    if sha(backup) != original_sha or sha(source) != original_sha:
        raise ValueError('Input changed during preparation')
    baseline = {'schema_version': 1, 'source': str(source), 'source_sha256': original_sha,
                'backup': str(backup), 'backup_sha256': original_sha,
                'school_template_sha256': template_sha,
                'style_spec_sha256': sha(ROOT / 'references' / 'style-spec.json')}
    save(work / 'baseline.json', baseline)
    for scope, name in [('styles', 'styles-inventory.json'), ('layout', 'layout-inventory.json'),
                        ('citations', 'citation-inventory.json')]:
        invoke({'action': 'inspect', 'scope': scope, 'source': str(backup), 'report': str(work / name)})
    invoke({'action': 'audit', 'source': str(backup), 'report': str(work / 'audit.json'),
            'markdown': str(work / 'audit.md')})
    audit = json.loads((work / 'audit.json').read_text(encoding='utf-8'))
    identified = audit.get('role_identification', {})
    save(work / 'style-map.draft.json', {
        'source_sha256': original_sha, 'reviewed': False, 'paragraphs': {}, 'tables': {},
        'suggestions_only': identified,
        'instruction': 'Read the inventories and paper. Copy only confirmed role assignments into paragraphs; classify tables separately. Never submit this draft as a reviewed map.'})
    save(work / 'layout-map.draft.json', {
        'source_sha256': original_sha, 'reviewed': False, 'title': {}, 'regions': [],
        'instruction': 'Confirm the actual title and front/body boundaries using the layout inventory. Read references/workflow.md. Re-inspect the exact input if an earlier stage changed the DOCX.'})
    save(work / 'citation-map.draft.json', {
        'schema_version': 1, 'source_sha256': original_sha, 'reviewed': False, 'spans': [],
        'instruction': 'Review citation-inventory.json and actual bibliography boundaries. Do not superscript formula/array/numeric markers. Re-inspect after previous edits.'})
    from review_plan import create_review_plan
    create_review_plan(backup, work / 'review', draft_path=work / 'style-map.draft.json')
    todo = '''# 下一步：先审核，不要直接生成“完成稿”

这里已经保留原稿并运行静态检查，尚未排版或查看打印页面。

1. 读取 skill 的 SKILL.md、references/format-requirements.md 和 references/workflow.md。
2. 读取 audit.md 及三份 inventory；按 review/index.json 逐批阅读 chunk，确认全文每段、每表及已有页眉脚/脚注，不能用首尾抽样代替全文。
3. 三个 *.draft.json 只是待审草稿；按 references/long-documents.md 填写各 chunk 的角色或逐项保留原因。未审记录保持未审，续跑从未审块继续。
4. 用 review_plan.py assemble 合并全部 chunk，生成正式 style-map 和 coverage.json，再对完整副本执行。不要把论文分成独立DOCX后重新拼接。每阶段保留输入/输出和报告，并重新绑定准确输入哈希。
5. 基础助手不重建学校封面、旧文自动编号或公式；不能改套一个新样式就宣称全部完成。未实现的项目须列清。
6. 渲染最终 DOCX，实际查看每一页。目录代码正确不等于缓存更新；TOC、显示标题编号、奇偶页眉都要看。
7. 运行 delivery_check.py，提供 --coverage，修改过的文档还必须提供 --chain；仅有静态结果或缺少证据时，交付措辞必须是“未验收稿/部分完成”，不能称为合格排版稿。

原稿基线是 baseline.json 和 original.docx，后续不能用成稿替换它们。
'''
    with (work / 'START.md').open('x', encoding='utf-8') as handle:
        handle.write(todo)
    if sha(source) != original_sha or sha(backup) != original_sha:
        raise ValueError('Source or backup changed; preparation invalid')
    result = {'operation_succeeded': True, 'delivery_ready': False, 'status': 'awaiting_mapping_review',
              'work_directory': str(work), 'source_unchanged': True, 'source_sha256': original_sha,
              'summary': audit.get('summary', {}),
              'next_read': str(work / 'START.md'),
              'review_index': str(work / 'review' / 'index.json'),
              'remaining_steps': ['Review mappings against the paper', 'Execute supported formatting changes',
                                  'Check unchanged research content', 'Render and review every final page', 'Run delivery_check']}
    save(work / 'prepare-result.json', result)
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--source', type=Path, required=True)
    parser.add_argument('--work-dir', type=Path, required=True)
    args = parser.parse_args()
    work_existed = args.work_dir.expanduser().resolve().exists()
    try:
        print(json.dumps(prepare(args.source, args.work_dir), ensure_ascii=False, indent=2))
        return 0
    except Exception as exc:
        failure = {'operation_succeeded': False, 'delivery_ready': False, 'status': 'preparation_failed', 'error': str(exc)}
        work = args.work_dir.expanduser().resolve()
        if not work_existed and work.is_dir():
            failure['work_directory'] = str(work)
            failure['partial_outputs'] = sorted(p.name for p in work.iterdir() if p.is_file())
            failure['remaining_steps'] = ['Inspect the failure; partial outputs are diagnostic only', 'Retry with a new directory after correcting the environment or input']
            result_path = work / 'prepare-result.json'
            if not result_path.exists():
                save(result_path, failure)
        print(json.dumps(failure, ensure_ascii=False))
        return 2


if __name__ == '__main__':
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    raise SystemExit(main())
