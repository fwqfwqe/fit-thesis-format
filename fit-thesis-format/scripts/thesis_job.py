#!/usr/bin/env python3
"""Persistent vendor-neutral thesis job. JSON in/out; no network or AI SDK.

A job drives deterministic edits and both Office renderers. The host must supply
actual semantic decisions and page observations; this module never invents them.
"""
from __future__ import annotations

import argparse
import copy
from contextlib import contextmanager
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import time
import uuid

SCRIPTS = Path(__file__).resolve().parent
ROOT = SCRIPTS.parent
ENGINES = ('word', 'wps')
PAGE_CHECKS = ('layout', 'headers_footers', 'numbering_toc', 'math_tables', 'content_order')


@contextmanager
def job_lock(work):
    """OS lock releases on crash; do not guess stale process IDs or delete locks."""
    lock = Path(work) / '.job.lock'
    with lock.open('a+b') as handle:
        if handle.tell() == 0:
            handle.write(b'0')
            handle.flush()
        handle.seek(0)
        try:
            if os.name == 'nt':
                import msvcrt
                msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError as exc:
            raise ValueError('This job is already executing in another process') from exc
        try:
            yield
        finally:
            handle.seek(0)
            if os.name == 'nt':
                msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(handle.fileno(), fcntl.LOCK_UN)


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def read(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))


def atomic_json(path, obj):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(path.name + '.' + uuid.uuid4().hex + '.tmp')
    try:
        with temp.open('x', encoding='utf-8') as out:
            json.dump(obj, out, ensure_ascii=False, indent=2)
            out.flush()
            os.fsync(out.fileno())
        os.replace(temp, path)
    finally:
        if temp.exists():
            temp.unlink()


def rules_fingerprint():
    files = sorted((ROOT / 'references').glob('*.json'))
    files += sorted((ROOT / 'assets').glob('*.docx'))
    # Code changes invalidate an existing job's execution evidence as well.
    files += sorted(SCRIPTS.glob('*.py')) + sorted(SCRIPTS.glob('*.ps1'))
    return {str(p.relative_to(ROOT)).replace('\\', '/'): sha(p) for p in files}


def bound_path(work, name):
    work = Path(work).resolve()
    path = (work / name).resolve()
    if path == work or work not in path.parents:
        raise ValueError('Task output escaped the task directory')
    return path


def save_state(work, state):
    state['updated_at'] = time.time()
    atomic_json(bound_path(work, 'job.json'), state)


def load_state(work):
    work = Path(work).resolve()
    state = read(bound_path(work, 'job.json'))
    if state.get('work_dir') != str(work):
        raise ValueError('Job directory moved; explicitly migrate and reverify evidence')
    source = Path(state['source'])
    if not source.is_file() or sha(source) != state['source_sha256']:
        raise ValueError('Original input changed; existing decisions and evidence are invalid')
    if sha(bound_path(work, 'original.docx')) != state['source_sha256']:
        raise ValueError('Immutable original copy changed')
    if state['rule_files'] != rules_fingerprint():
        raise ValueError('Rules or code changed; restart from original in a new job directory')
    return state


def doctor(work_dir=None, smoke=True):
    import office_probe
    return office_probe.doctor(work_dir=work_dir, smoke=smoke)


def new_job(request):
    source = Path(request['source']).expanduser().resolve(strict=True)
    if source.suffix.lower() != '.docx' or not source.is_file():
        raise ValueError('Source must be an existing DOCX')
    work = Path(request['work_dir']).expanduser().resolve()
    if work.exists() or work.is_symlink() or work == source or source in work.parents:
        raise ValueError('work_dir must be a new directory distinct from the input')
    # Do not create an arbitrary caller-specified final path; all artifacts live here.
    work.mkdir(parents=True, exist_ok=False)
    source_hash = sha(source)
    shutil.copyfile(source, bound_path(work, 'original.docx'))
    state = {'schema_version': 10, 'job_id': uuid.uuid4().hex, 'work_dir': str(work),
             'source': str(source), 'source_sha256': source_hash,
             'rule_files': rules_fingerprint(), 'profile': request.get('profile'),
             'decisions': request.get('decisions'), 'stages': [], 'iterations': [],
             'renders': {}, 'reviews': {}, 'status': 'created', 'revision': 0,
             'required_engines': list(ENGINES), 'delivery_ready': False}
    if isinstance(state['decisions'], str):
        state['decisions'] = read(state['decisions'])
    if isinstance(state['profile'], str):
        state['profile'] = read(state['profile'])
    save_state(work, state)
    with job_lock(work):
        return advance(work, state)


def mapping_from(result):
    value = result.get('final_mapping') or result.get('mapping') or result.get('role_map') or result.get('map')
    if isinstance(value, str):
        value = read(value)
    return value if isinstance(value, dict) else None


def rebind_final_mapping(before, final, mapping):
    """Retain reviewed choices across a verified layout-cache-only refresh.

    TOC cache paragraphs may be replaced, so global paragraph IDs can shift.
    Every other paragraph is paired in order and independently compared; table
    settings are copied in full, never inferred again from their appearance.
    """
    import thesis_format as F
    from field_cache_merge import spans, pspan
    from format_audit import text as audit_text
    from format_engine import inspect_document
    from integrity import story_tokens
    verify_mapping_complete(before, mapping)
    ap, bp = F.package(before), F.package(final)
    ar, br = F.parse(ap['word/document.xml']), F.parse(bp['word/document.xml'])
    aps, ats = F.indexed(ar); bps, bts = F.indexed(br)
    def toc_nodes(root):
        excluded = set()
        for field in spans(root):
            if field['kind'] != 'TOC':
                continue
            first, last = pspan(field)
            if first is None or last is None or first.getparent() is not last.getparent():
                raise ValueError('TOC cache has incompatible paragraph boundaries')
            parent = first.getparent()
            for block in list(parent)[parent.index(first):parent.index(last)+1]:
                if block.tag != F.q('p'):
                    raise ValueError('TOC cache contains non-paragraph content')
                excluded.update(block.iter(F.q('p')))
        return excluded
    ae, be = toc_nodes(ar), toc_nodes(br)
    aordered = [(pid, p) for pid, p in aps.items() if p not in ae]
    bordered = [(pid, p) for pid, p in bps.items() if p not in be]
    if len(aordered) != len(bordered) or len(ats) != len(bts):
        raise ValueError('Non-TOC object count changed during field refresh')
    rebound = copy.deepcopy(mapping)
    rebound.update(source_sha256=sha(final), paragraphs={}, tables={}, preserve={}, story_paragraphs={})
    for (aid, a), (bid, b) in zip(aordered, bordered):
        if story_tokens(a, ap, 'word/document.xml') != story_tokens(b, bp, 'word/document.xml'):
            raise ValueError('Non-TOC paragraph changed during field refresh: ' + aid + ' -> ' + bid)
        if aid in mapping.get('paragraphs', {}):
            rebound['paragraphs'][bid] = {**copy.deepcopy(mapping['paragraphs'][aid]), 'expected_text': audit_text(b)}
        if aid in mapping.get('preserve', {}):
            obj = next((x for x in b.iterancestors() if x.tag in F.PROTECTED_TAGS), None)
            if obj is None:
                raise ValueError('Protected object lost its enclosing structure')
            rebound['preserve'][bid] = {**copy.deepcopy(mapping['preserve'][aid]), 'expected_text': audit_text(b),
                                       'object_sha256': F.sha(F.xml(obj))}
    # The sole fresh classification is for generated TOC cache paragraphs.
    fresh = mapping_from(inspect_document(final))
    for pid, p in bps.items():
        if p not in be:
            continue
        # Word can emit a one-point, empty field-end carrier. It is not a
        # displayed TOC entry and has no content to classify as a TOC style.
        if not audit_text(p).strip() and not F.ptext(p).strip():
            continue
        entry = fresh.get('paragraphs', {}).get(pid, {})
        if entry.get('role') not in ('toc1', 'toc2'):
            if audit_text(p).strip():
                raise ValueError('Generated TOC paragraph lacks a known level: ' + pid)
            entry = {'role': 'toc1'}  # empty field boundary, known TOC region
        rebound['paragraphs'][pid] = {**entry, 'expected_text': audit_text(p)}
    for (aid, a), (bid, b) in zip(ats.items(), bts.items()):
        if story_tokens(a, ap, 'word/document.xml') != story_tokens(b, bp, 'word/document.xml'):
            raise ValueError('Table identity changed during field refresh: ' + aid)
        rebound['tables'][bid] = copy.deepcopy(mapping['tables'][aid])
    if {n for n in ap if F.allowed_story(n)} != {n for n in bp if F.allowed_story(n)}:
        raise ValueError('Story part identities changed during cache-only refresh')
    for name in ap:
        if not F.allowed_story(name):
            continue
        asp, _ = F.indexed(F.parse(ap[name])); bsp, _ = F.indexed(F.parse(bp[name]))
        if len(asp) != len(bsp):
            raise ValueError('Story paragraph count changed: ' + name)
        selected = {}
        for (aid, a), (bid, b) in zip(asp.items(), bsp.items()):
            if story_tokens(a, ap, name) != story_tokens(b, bp, name):
                raise ValueError('Story semantics changed: ' + name + ':' + aid)
            if aid in mapping.get('story_paragraphs', {}).get(name, {}):
                selected[bid] = {**copy.deepcopy(mapping['story_paragraphs'][name][aid]), 'expected_text': audit_text(b)}
        if selected:
            rebound['story_paragraphs'][name] = selected
    verify_mapping_complete(final, rebound)
    return rebound


def verify_mapping_complete(final, mapping):
    """Do not silently skip a table or story because classification was lost."""
    from thesis_format import package, parse, indexed, q, ptext, allowed_story, PROTECTED_TAGS, xml
    from format_audit import text as audit_text
    if not mapping or mapping.get('source_sha256') != sha(final):
        raise ValueError('Final semantic map must be bound to the final DOCX')
    parts = package(final)
    ps, ts = indexed(parse(parts['word/document.xml']))
    preserved = mapping.get('preserve', {})
    for pid, decision in preserved.items():
        if pid not in ps or not isinstance(decision, dict):
            raise ValueError('Invalid protected-object coverage entry')
        parent = next((a for a in ps[pid].iterancestors() if a.tag in PROTECTED_TAGS), None)
        if parent is None or decision.get('expected_text') != audit_text(ps[pid]):
            raise ValueError('Preserve coverage must identify exact text nested inside a protected object')
        if decision.get('object_sha256') != hashlib.sha256(xml(parent)).hexdigest():
            raise ValueError('Protected object changed after semantic review')
        if not str(decision.get('reason', '')).strip() or decision.get('pending_visual') is not True:
            raise ValueError('Protected object needs a reason and the mandatory visual review gate')
    missing = [pid for pid, p in ps.items()
               if (ptext(p).strip() or p.find('.//' + q('drawing')) is not None
                   or p.xpath('.//*[local-name()="oMath" or local-name()="object"]'))
               and pid not in mapping.get('paragraphs', {}) and pid not in preserved]
    missing += [tid for tid in ts if tid not in mapping.get('tables', {})]
    for name, data in parts.items():
        if not allowed_story(name):
            continue
        sp, _ = indexed(parse(data))
        selected = mapping.get('story_paragraphs', {}).get(name, {})
        missing += [name + ':' + pid for pid, p in sp.items() if ptext(p).strip() and pid not in selected]
    if missing:
        raise ValueError('Unclassified final objects: ' + ', '.join(missing[:30]))
    return {'passed': True, 'paragraphs': len(ps), 'tables': len(ts)}


def verify_stage(source, final, kind, report):
    from integrity import verify_transition
    result = verify_transition(source, final, kind, report=report)
    if result.get('passed') is not True:
        raise ValueError('Content protection failed: ' + json.dumps(result.get('errors'), ensure_ascii=False))
    return result


def pages_for(pdf, directory):
    completed = subprocess.run([sys.executable, str(SCRIPTS / 'render_pdf.py'), str(pdf), str(directory),
                                '--dpi', '144'], capture_output=True, text=True, encoding='utf-8', timeout=240,
                               env=dict(os.environ, PYTHONUTF8='1', PYTHONIOENCODING='utf-8'))
    if completed.returncode:
        raise ValueError('Page rendering failed: ' + (completed.stderr or completed.stdout)[-1500:])
    return Path(directory) / 'manifest.json'


def advance(work, state=None):
    """Advance until an actual host judgment or unavailable capability is needed."""
    work = Path(work).resolve()
    state = load_state(work) if state is None else state
    if state.get('status') == 'no_progress':
        return response(work, state)
    try:
        if 'capabilities' not in state:
            state['capabilities'] = doctor(bound_path(work, 'capabilities'), smoke=True)
            save_state(work, state)
        if 'candidate' not in state:
            import format_engine
            revision = state['revision']
            attempt = bound_path(work, f'attempt-{revision:03d}-{uuid.uuid4().hex[:8]}')
            attempt.mkdir()
            candidate = attempt / 'formatted.docx'
            source = bound_path(work, 'original.docx')
            if not state.get('inspection'):
                state['inspection'] = format_engine.inspect_document(source)
                atomic_json(bound_path(work,'inspection.json'),state['inspection'])
            state['active_attempt'] = str(attempt)
            state['status'] = 'formatting'
            save_state(work, state)
            result = format_engine.format_document(source, candidate, decisions=state.get('decisions'), profile=state.get('profile'))
            result['decisions'], result['profile'] = state.get('decisions'), state.get('profile')
            atomic_json(attempt / 'format-report.json', result)
            unresolved = result.get('unresolved', result.get('unresolved_issues', result.get('issues', [])))
            state['content_advice'] = result.get('content_advice',[]) + [item for item in unresolved if item.get('code') == 'missing_identity_fields']
            unresolved = [item for item in unresolved if item.get('code') != 'missing_identity_fields']
            if not candidate.is_file():
                state['status'] = 'semantic_review_required'
                state['blocking_items'] = unresolved or [{'reason': 'Formatting requires complete semantic decisions', 'inspection': result}]
                save_state(work, state)
                return response(work, state)
            protection = verify_stage(source, candidate, 'format', result)
            atomic_json(attempt / 'integrity-format.json', protection)
            state['stages'] = [{'kind': 'format', 'source': str(source), 'output': str(candidate),
                                'source_sha256': sha(source), 'output_sha256': sha(candidate),
                                'report': str(attempt / 'format-report.json')}]
            state['format_report'] = str(attempt / 'format-report.json')
            state['candidate'] = str(candidate)
            state['unresolved'] = unresolved
            state['mapping'] = mapping_from(result)
            from format_audit import Audit
            # Inspect once BEFORE slow Office work so the host can batch
            # structural corrections rather than render a known broken draft.
            audit = Audit(candidate, mapping=state['mapping']).run()
            path = attempt / 'static-audit.json'
            atomic_json(path, audit)
            state['static_audit'] = str(path)
            static = [i for i in audit['issues'] if i.get('status') == 'auto_fixable'
                      and i.get('severity') in ('warning','error')
                      and not i.get('category','').startswith(('toc.','fields.'))]
            state['unresolved'] += static
            state['status'] = 'formatted'
            save_state(work, state)
        if state.get('unresolved'):
            state['status'] = 'semantic_review_required'
            state['blocking_items'] = state['unresolved']
            save_state(work, state)
            return response(work, state)
        if not state['capabilities'].get('dual_render_capable'):
            state['status'] = 'capability_required'
            state['blocking_items'] = [{'check': 'dual_engine', 'reason': 'Candidate and static audit prepared; both Word and WPS must pass actual rendering before delivery'}]
            save_state(work, state)
            return response(work, state)
        if 'final' not in state:
            import office_probe
            attempt = Path(state['active_attempt'])
            final = attempt / ('fields-' + uuid.uuid4().hex[:8] + '.docx')
            render = office_probe.render('word', state['candidate'], attempt / (final.stem + '.pdf'), saved_docx=final)
            field_report = attempt / (final.stem + '.json')
            atomic_json(field_report, render)
            protection = verify_stage(state['candidate'], final, 'office_fields', render)
            atomic_json(attempt / (final.stem + '-integrity.json'), protection)
            state['stages'].append({'kind': 'office_fields', 'source': state['candidate'], 'output': str(final),
                                    'source_sha256': sha(state['candidate']), 'output_sha256': sha(final), 'report': str(field_report)})
            state['mapping'] = rebind_final_mapping(state['candidate'], final, state['mapping'])
            state['final'] = str(final)
            state['final_sha256'] = sha(final)
            state['status'] = 'fields_refreshed'
            save_state(work, state)
        if 'final_mapping' not in state:
            verify_mapping_complete(state['final'], state.get('mapping'))
            path = Path(state['active_attempt']) / 'final-mapping.json'
            atomic_json(path, state['mapping'])
            state['final_mapping'] = str(path)
            save_state(work, state)
        for engine in ENGINES:
            if engine in state['renders']:
                continue
            import office_probe
            directory = Path(state['active_attempt']) / (engine + '-' + uuid.uuid4().hex[:8])
            directory.mkdir()
            rendering = office_probe.render(engine, state['final'], directory / 'document.pdf')
            report_path = directory / 'render.json'
            atomic_json(report_path, rendering)
            manifest_path = pages_for(directory / 'document.pdf', directory / 'pages')
            manifest = read(manifest_path)
            from delivery_check import continuation_requirements
            continued = continuation_requirements(rendering, state['final'])
            review = {'final_docx_sha256': sha(state['final']), 'pdf_sha256': rendering['pdf_sha256'],
                      'engine_key': engine, 'reviewer': '', 'pages': [
                          {'page': item['page'], 'image_sha256': item['sha256'], 'status': 'pending',
                           'checks': {k: 'pending' for k in PAGE_CHECKS}, 'notes': '',
                           'continued_tables': [{'table_index': i, 'caption_label': '', 'header_repeated': False,
                                                 'no_missing_rows': False} for i in continued.get(item['page'], [])]}
                          for item in manifest['pages']]}
            atomic_json(directory / 'review.pending.json', review)
            state['renders'][engine] = {'render_manifest': str(report_path), 'page_manifest': str(manifest_path),
                                        'review_template': str(directory / 'review.pending.json')}
            state['status'] = 'rendering'
            save_state(work, state)
        if not all(engine in state['reviews'] for engine in ENGINES) or any(
                any(p.get('status') == 'pending' for p in read(path)['pages']) for path in state['reviews'].values()):
            state['status'] = 'visual_review_required'
            save_state(work, state)
            return response(work, state)
        result = verify_job(work, state)
        state['verification'] = result
        state['delivery_ready'] = result['delivery_ready']
        state['status'] = 'complete' if result['delivery_ready'] else 'correction_required'
        save_state(work, state)
        return response(work, state)
    except Exception as exc:
        state['status'] = 'execution_failed'
        state['delivery_ready'] = False
        state['last_error'] = str(exc)
        save_state(work, state)
        return response(work, state)


def verify_job(work, state=None):
    from delivery_check import check_dual_render
    from format_audit import Audit
    state = load_state(work) if state is None else state
    # Even callers holding a state object must recheck the immutable bindings.
    load_state(work)
    failures, details = [], {}
    final = Path(state.get('final', ''))
    if not final.is_file() or sha(final) != state.get('final_sha256'):
        return {'delivery_ready': False, 'blocking_items': [{'check': 'final', 'reason': 'Final file is missing or changed'}]}
    try:
        previous = state['source_sha256']
        for stage in state['stages']:
            if sha(stage['source']) != previous or sha(stage['output']) != stage['output_sha256']:
                raise ValueError('Disconnected or changed transformation stage')
            verify_stage(stage['source'], stage['output'], stage['kind'], read(stage['report']))
            previous = sha(stage['output'])
        if previous != sha(final):
            raise ValueError('Final is outside the verified transformation chain')
        details['content_protection'] = 'passed_fresh_verification'
    except Exception as exc:
        failures.append({'check': 'content_protection', 'reason': str(exc)})
    try:
        mapping = read(state['final_mapping'])
        details['coverage'] = verify_mapping_complete(final, mapping)
        audit = Audit(final, mapping=mapping).run()
        unresolved = [i for i in audit['issues'] if i.get('severity') in ('warning', 'error')]
        details['audit'] = audit
        if unresolved or state.get('unresolved'):
            failures.append({'check': 'format', 'reason': 'Unresolved applicable format rules', 'issues': unresolved + state.get('unresolved', [])})
    except Exception as exc:
        failures.append({'check': 'final_mapping_and_audit', 'reason': str(exc)})
    try:
        evidence = {engine: {**state['renders'].get(engine, {}), 'review': state['reviews'].get(engine)} for engine in ENGINES}
        details['dual_render'] = check_dual_render(final, evidence)
    except Exception as exc:
        failures.append({'check': 'dual_render', 'reason': str(exc)})
    try:
        from rule_coverage import evaluate
        coverage = evaluate(final, read(state['final_mapping']), details.get('audit', {}),
                            details.get('content_protection') == 'passed_fresh_verification', details.get('dual_render'))
        details['rule_results'] = coverage['rule_results']
        details['content_advice'] = state.get('content_advice', [])
        if not coverage['passed']:
            failures.append({'check': 'rule_coverage', 'reason': 'Every applicable rule must have positive evidence',
                             'rules': [r['rule_id'] for r in coverage['rule_results'] if r['status'] == 'fail']})
    except Exception as exc:
        failures.append({'check': 'rule_coverage', 'reason': str(exc)})
    return {'delivery_ready': not failures, 'source_sha256': state['source_sha256'], 'final_docx_sha256': sha(final),
            'blocking_items': failures, 'details': details}


def response(work, state):
    from job_guidance import repair_plan
    # Keep stdout useful to small-context hosts. Full evidence stays in job.json
    # and inspection.json; sampling here is transport, never review coverage.
    def sample(items, limit, budget):
        selected=[];used=0
        for item in items:
            size=len(json.dumps(item,ensure_ascii=False).encode('utf-8'))
            if len(selected)<limit and used+size<=budget:
                selected.append(item);used+=size
        return selected
    plan=repair_plan(state,work);remaining=16384;group_omitted={}
    for group,items in list(plan['groups'].items()):
        selected=sample(items,10,remaining)
        remaining-=len(json.dumps(selected,ensure_ascii=False).encode('utf-8'))
        plan['groups'][group]=selected;group_omitted[group]=len(items)-len(selected)
    plan['sample_omitted_counts']=group_omitted
    plan['complete_findings']={'file':str(Path(work)/'job.json'),
        'json_pointers':['/verification/blocking_items','/blocking_items','/unresolved','/content_advice'],
        'complete_review_required':True}
    blockers=state.get('verification',{}).get('blocking_items',state.get('blocking_items',[]))
    selected_blockers=sample(blockers,30,8192)
    return {'protocol_version': 10, 'ok': state['status'] != 'execution_failed',
            'operation_succeeded': state['status'] != 'execution_failed',
            'work_dir': str(work), 'status': state['status'], 'delivery_ready': state.get('delivery_ready', False),
            'candidate': state.get('final', state.get('candidate')), 'next': next_tasks(work, state),
            'repair_plan':plan,
            'blocking_items':selected_blockers,'blocking_item_count':len(blockers),
            'blocking_items_omitted':len(blockers)-len(selected_blockers),
            'blocking_items_file':{'file':str(Path(work)/'job.json'),
                'json_pointer':'/verification/blocking_items' if 'verification' in state else '/blocking_items'},
            'error': state.get('last_error')}


def next_tasks(work, state=None):
    state = load_state(work) if state is None else state
    if state['status'] == 'visual_review_required':
        tasks = []
        for engine in ENGINES:
            saved = state['reviews'].get(engine)
            pending = [p['page'] for p in read(saved)['pages'] if p.get('status') == 'pending'] if saved else None
            if saved and not pending:
                continue
            tasks.append({'kind': 'page_review', 'engine': engine, **state['renders'][engine],
                          'saved_review': saved, 'pending_pages': pending, 'final_docx_sha256': state['final_sha256']})
        return tasks
    if state['status'] in ('semantic_review_required', 'correction_required'):
        from job_guidance import repair_plan
        inspection = state.get('inspection') or {}
        decided=state.get('decisions') or {}
        confirmed=set(decided.get('paragraphs',{}))|set(decided.get('preserve',{}))
        candidates=inspection.get('candidates',[])
        pending=[c for c in candidates if c.get('id') not in confirmed]
        sample=[];used=0
        for c in pending:
            item={key:c[key] for key in ('id','role','text','reason') if key in c}
            size=len(json.dumps(item,ensure_ascii=False).encode('utf-8'))
            if len(sample)<30 and used+size<=16384:
                sample.append(item);used+=size
        summary={'source_sha256':state['source_sha256'],'counts':inspection.get('counts',{}),
                 'candidate_count':len(candidates),'pending_candidate_count':len(pending),
                 'candidates':sample,'sample_omitted_count':len(pending)-len(sample),
                 'complete_review_required':True,'sample_is_not_full_coverage':True}
        return [{'kind': 'decisions', 'source_sha256': state['source_sha256'],
                 'inspection_file':str(Path(work)/'inspection.json'),
                 'inspection':summary,
                 'decision_template': {'source_sha256': state['source_sha256'], 'paragraphs': {}},
                 'completion_policy': repair_plan(state, work)['completion_policy'],
                 'decision_template_is_example_only': True,
                 'existing_decisions': {'file': str(Path(work) / 'job.json'), 'json_pointer': '/decisions',
                                        'merge_before_submit': True},
                 'evidence': {'format_report': state.get('format_report'),
                              'renders': state.get('renders', {}), 'reviews': state.get('reviews', {})},
                 'submit_request': {'action': 'submit', 'work_dir': str(work), 'kind': 'decisions',
                                    'decisions': '<absolute-path-to-reviewed-decisions.json>'},
                 'issues_file':{'file':str(Path(work)/'job.json'),
                    'json_pointer':'/verification/blocking_items' if 'verification' in state else '/blocking_items'}}]
    if state['status'] == 'execution_failed':
        return [{'kind': 'repair_environment_or_execution', 'reason': state.get('last_error')}]
    if state['status'] == 'capability_required':
        return [{'kind': 'capabilities', 'required_engines': list(ENGINES), 'actual': state.get('capabilities')}]
    if state['status'] == 'no_progress':
        return [{'kind': 'diagnose_repeated_failure',
                 'issues': state.get('blocking_items', []),
                 'instruction': 'Diagnose the unchanged failure; submit materially corrected source-bound decisions. Do not repeat resume or claim completion.'}]
    return []


def submit(work, request):
    state = load_state(work)
    if request.get('kind') == 'review':
        engine = request.get('engine')
        if engine not in ENGINES or engine not in state['renders']:
            raise ValueError('Unknown or not-yet-rendered engine')
        review = request.get('review')
        if isinstance(review, str):
            review = read(review)
        if not isinstance(review, dict) or review.get('final_docx_sha256') != state.get('final_sha256'):
            raise ValueError('Review must be bound to the actual final DOCX')
        path = Path(state['active_attempt']) / (engine + '-review-' + uuid.uuid4().hex[:8] + '.json')
        atomic_json(path, review)
        from delivery_check import check_render
        proof = state['renders'][engine]
        check_render(Path(state['final']), Path(proof['render_manifest']), Path(proof['page_manifest']), path, allow_incomplete=True)
        state['reviews'][engine] = str(path)
    elif request.get('kind') == 'decisions':
        decisions = request.get('decisions')
        if isinstance(decisions, str):
            decisions = read(decisions)
        if not isinstance(decisions, dict) or decisions.get('source_sha256') != state['source_sha256']:
            raise ValueError('Semantic decisions must be bound to the original source')
        signature = hashlib.sha256(json.dumps(decisions, sort_keys=True, ensure_ascii=False).encode()).hexdigest()
        repeats = state.get('unchanged_submissions', 0) + 1 if state.get('decision_signature') == signature else 0
        state['unchanged_submissions'] = repeats
        if repeats >= 2:
            state['status'] = 'no_progress'
            state['blocking_items'] = [{'reason': 'Two repeated identical correction submissions; diagnose the unresolved issue instead of retrying unchanged work'}]
        else:
            state['iterations'].append({k: state.get(k) for k in ('revision', 'candidate', 'final', 'stages', 'verification', 'renders', 'reviews')})
            for key in ('candidate', 'final', 'final_sha256', 'format_report', 'final_mapping', 'verification', 'last_error', 'blocking_items', 'unresolved'):
                state.pop(key, None)
            state['decisions'], state['decision_signature'] = decisions, signature
            state['revision'] += 1
            state['renders'], state['reviews'], state['stages'] = {}, {}, []
            state['delivery_ready'], state['status'] = False, 'created'
    else:
        raise ValueError('submit kind must be decisions or review')
    save_state(work, state)
    return advance(work)


def _execute(request):
    if not isinstance(request, dict):
        raise ValueError('Request must be an object')
    action = request.get('action')
    if action == 'doctor':
        return {'ok': True, 'delivery_ready': False, 'status': 'capabilities_checked',
                'result': doctor(request.get('work_dir'), smoke=request.get('smoke', True))}
    if action == 'run':
        return new_job(request)
    work = Path(request['work_dir']).expanduser().resolve()
    if action == 'next':
        return response(work, load_state(work))
    if action == 'submit':
        return submit(work, request)
    if action == 'resume':
        state = load_state(work)
        if state['status'] == 'capability_required' or request.get('reprobe') is True:
            state.pop('capabilities', None)
        state.pop('last_error', None)
        save_state(work, state)
        return advance(work)
    if action == 'verify':
        state = load_state(work)
        result = verify_job(work, state)
        state['verification'], state['delivery_ready'] = result, result['delivery_ready']
        state['status'] = 'complete' if result['delivery_ready'] else 'correction_required'
        save_state(work, state)
        return response(work, state)
    raise ValueError('Unknown v10 task action')


def execute(request):
    if not isinstance(request, dict):
        raise ValueError('Request must be a JSON object')
    specs = {'doctor': (set(), {'work_dir','smoke','legacy'}),
             'run': ({'source','work_dir'}, {'profile','decisions'}),
             'next': ({'work_dir'}, set()), 'resume': ({'work_dir'}, {'reprobe'}),
             'verify': ({'work_dir'}, set()),
             'submit': ({'work_dir','kind'}, {'engine','review','decisions'})}
    action = request.get('action')
    if action not in specs:
        raise ValueError('Unknown job action')
    required, optional = specs[action]
    if required - request.keys() or request.keys() - required - optional - {'action'}:
        raise ValueError('Missing or unexpected task parameters')
    if isinstance(request, dict) and request.get('action') in ('resume', 'submit', 'verify'):
        work = Path(request['work_dir']).expanduser().resolve(strict=True)
        with job_lock(work):
            return _execute(request)
    return _execute(request)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--request', required=True)
    args = parser.parse_args()
    try:
        request = json.loads(sys.stdin.read()) if args.request == '-' else read(args.request)
        result = execute(request)
    except Exception as exc:
        result = {'ok': False, 'operation_succeeded': False, 'delivery_ready': False, 'status': 'failed', 'error': str(exc)}
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get('ok') else 2


if __name__ == '__main__':
    for stream in (sys.stdin, sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            stream.reconfigure(encoding='utf-8')
    raise SystemExit(main())
