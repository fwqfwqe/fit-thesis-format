"""Host-neutral recovery instructions; never alter verification or grant a pass.

This is routing, not a second validator. Unknown findings remain actionable,
and content gaps cannot hide independent formatting failures.
"""
from __future__ import annotations

import json


GROUPS = ('format_repairs', 'evidence_checks', 'missing_content',
          'capability_repairs', 'execution_repairs')
MISSING_CONTENT = {
    'missing_identity_fields', 'missing_english_title', 'front.english_title_missing',
    'front.en_title_missing', 'front.missing_english_title', 'missing_acknowledgement',
}
EVIDENCE_CHECKS = {
    'citation.missing_year', 'reference.author_listing_needs_review',
    'structure.acknowledgement_missing_or_unconfirmed', 'front.required_part_missing',
    'role_confirmation', 'legacy_equation_editability', 'equation_label_ambiguous',
    'equation_number_missing', 'broken_ref_ambiguous', 'numeric_citation_mixed_run',
    'citation_in_revision', 'coverage.unclassified_content',
}
FORMAT_PREFIXES = ('paragraph.', 'style.', 'heading.', 'toc.', 'pagination.',
                   'header.', 'header_footer.', 'page.', 'table.', 'math.',
                   'fields.', 'front.', 'structure.', 'citation.number_position')


def _leaves(items):
    """Unwrap verification containers without losing unknown leaf findings."""
    for item in items or []:
        if not isinstance(item, dict):
            yield {'reason': str(item)}
            continue
        nested = item.get('issues')
        if isinstance(nested, list) and nested:
            yield from _leaves(nested)
        else:
            yield item


def _route(item):
    code = item.get('code') or item.get('category') or item.get('check') or ''
    if code in MISSING_CONTENT:
        return 'missing_content', '保留原稿及正常填写空位；先完成其余格式，集中报告确实缺失的内容。'
    if code in EVIDENCE_CHECKS or item.get('status') in ('requires_role_review', 'requires_content_authorization'):
        return 'evidence_checks', '先检查原稿上下文和本稿书目；有唯一依据就提交来源绑定修复，确无依据才列为待补。'
    if code in ('dual_engine', 'capabilities'):
        return 'capability_repairs', '完成可执行的格式和静态审计；补齐真实渲染能力后 resume，不能把其他引擎冒充 Word/WPS。'
    if code in ('execution', 'content_protection', 'final', 'final_mapping_and_audit'):
        return 'execution_repairs', '查看具体错误并修复执行或证据链；保留原稿和失败证据，不能重设基线。'
    if code.startswith(FORMAT_PREFIXES) or item.get('status') == 'auto_fixable':
        return 'format_repairs', '结合具体定位修复后 submit decisions；同类问题批量处理，再验证实际页。'
    if code == 'dual_render':
        return 'format_repairs', '打开各引擎 render/review 的失败页面和说明；修正分页、版式或证据，不重复提交相同复核。'
    return 'evidence_checks', '诊断这条未解决问题并确定修复；不能忽略未知问题或直接交付候选稿。'


def repair_plan(state, work_dir):
    """Return compact executable follow-up and grouped original findings."""
    verification = state.get('verification') or {}
    findings = verification.get('blocking_items', state.get('blocking_items', []))
    findings = list(_leaves(findings)) + list(_leaves(state.get('unresolved', [])))
    findings += list(_leaves(state.get('content_advice', [])))
    groups = {name: [] for name in GROUPS}
    seen = set()
    for item in findings:
        if item.get('severity') == 'info':
            continue
        signature = json.dumps(item, sort_keys=True, ensure_ascii=False, default=str)
        if signature in seen:
            continue
        seen.add(signature)
        group, action = _route(item)
        groups[group].append({'finding': item, 'action': action})
    ready = state.get('delivery_ready') is True and state.get('status') == 'complete'
    status = state.get('status', 'created')
    must_repair = bool(groups['format_repairs'] or groups['evidence_checks'])
    if ready:
        next_action = 'deliver_verified_docx'
    elif status == 'no_progress':
        next_action = 'diagnose_repeated_failure'
    elif status == 'execution_failed':
        next_action = 'repair_execution_then_resume'
    elif status == 'capability_required':
        next_action = 'complete_available_work_then_restore_capability'
    elif status == 'visual_review_required':
        next_action = 'inspect_pending_page_images_then_submit_review'
    elif must_repair:
        next_action = 'repair_and_reverify'
    elif groups['missing_content']:
        next_action = 'report_content_gaps_after_available_repairs'
    else:
        next_action = 'follow_next_task'
    return {
        'guidance_version': 1,
        'source_sha256': state.get('source_sha256'),
        'action': next_action,
        'completion_policy': {
            'may_claim_complete': ready,
            'must_continue_format_repairs': must_repair and not ready,
            'missing_content_does_not_waive_formatting': True,
            'must_not_invent_missing_content': True,
            'must_not_repeat_unchanged_submission': True,
        },
        'counts': {name: len(values) for name, values in groups.items()},
        'groups': {name: values for name, values in groups.items() if values},
        'commands': {
            'next': {'action': 'next', 'work_dir': str(work_dir)},
            'resume': {'action': 'resume', 'work_dir': str(work_dir)},
            'verify': {'action': 'verify', 'work_dir': str(work_dir)},
        },
        'repair_reference': 'references/layout-repairs.md',
        'host_reference': 'references/host-adapters.md',
    }
