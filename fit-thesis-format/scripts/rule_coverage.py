"""Positive evidence coverage of every enabled v10 rule; no silent exemptions."""
from pathlib import Path
import json
import re

import thesis_format as F

ROOT = Path(__file__).resolve().parent.parent


def evaluate(final, mapping, audit, protection_passed, dual_evidence=None):
    registry = json.loads((ROOT / 'references/rules-v10.json').read_text(encoding='utf-8'))
    parts = F.package(final); root = F.parse(parts['word/document.xml'])
    ps, ts = F.indexed(root); body = root.find(F.q('body'))
    roles = {pid: v if isinstance(v, str) else v.get('role') for pid, v in mapping.get('paragraphs', {}).items()}
    tables = mapping.get('tables', {})
    byrole = {}
    for pid, role in roles.items():byrole.setdefault(role, []).append(ps[pid])
    issues = [x for x in audit.get('issues', []) if x.get('severity') in ('warning','error')]
    final_hash = F.sha(Path(final).read_bytes())
    audit_bound = audit.get('source_sha256') == final_hash and audit.get('source_unchanged') is True
    dual = bool(dual_evidence and all(dual_evidence.get(e, {}).get('all_pages_passed') is True for e in ('word','wps'))
                and dual_evidence.get('same_final_docx_sha256') == F.sha(Path(final).read_bytes()))
    def has(*role_names):return any(byrole.get(r) for r in role_names)
    def clean(*prefixes):return audit_bound and not any(any(x.get('category','').startswith(p) for p in prefixes) for x in issues)
    def val(el, attr):return el.get(F.q(attr)) if el is not None else None
    def texts(el):return ''.join(el.xpath('.//w:t/text()', namespaces=F.NS))
    fulltext = texts(root)
    front_roles = ('cover_banner','commitment_heading','cn_abstract_heading','en_abstract_heading','cn_keywords','en_keywords','toc_heading')
    from format_audit import Audit
    cover_ps = [p for pid,p in ps.items() if roles.get(pid, '').startswith('cover_')]
    native_cover_logo = Audit(final, mapping=mapping).native_cover_media_match(cover_ps)
    from format_engine import front_template_geometry_issues
    front_geometry_issues = front_template_geometry_issues(root,F.parse(parts['word/styles.xml']))
    codes = [''.join(p.xpath('.//w:instrText/text()', namespaces=F.NS)) for p in ps.values()]
    toc_code = next((c for c in codes if re.search(r'\bTOC\b', c)), '')
    sections = root.findall('.//w:sectPr', F.NS)
    spec = F.SPEC['page']
    geometry = bool(sections)
    for sec in sections:
        for node, attrs in (('pgSz', {'w':'width_twips','h':'height_twips'}),
                            ('pgMar', {'top':'top_twips','bottom':'bottom_twips','left':'left_twips','right':'right_twips','header':'header_twips','footer':'footer_twips'})):
            el = sec.find(F.q(node))
            geometry = geometry and all(val(el, a) == str(spec[k]) for a,k in attrs.items())
    heading_ps = sum((byrole.get(r, []) for r in ('chapter','section','subsection','subsubsection')), [])
    heading_native = bool(byrole.get('chapter')) and 'word/numbering.xml' in parts and all(
        p.find('w:pPr/w:numPr/w:numId', F.NS) is not None and val(p.find('w:pPr/w:numPr/w:numId', F.NS), 'val') != '0' for p in heading_ps)
    style_ids = {x.get(F.q('styleId')) for x in F.parse(parts['word/styles.xml']).findall(F.q('style'))}
    style_bindings = bool(roles) and all(val(p.find('w:pPr/w:pStyle', F.NS), 'val') in style_ids for pid,p in ps.items() if pid in roles)
    target_roles = ['cn_abstract_heading','en_abstract_heading','chapter','section','reference_heading','appendix_heading','acknowledgement_heading']
    toc_targets = bool(toc_code) and all(F.SPEC['roles'][r]['name'] + ',' + ('2' if r=='section' else '1') in toc_code
                                       for r in target_roles if byrole.get(r))
    toc_levels = bool(toc_code) and not re.search(r',\s*[3-9](?:[,"\s]|$)', toc_code) and not re.search(r'\\o\s+"?1-[3-9]', toc_code)
    page_formats = {val(s.find(F.q('pgNumType')), 'fmt') for s in sections}
    numbering_sections = len(sections) >= 3 and {'upperRoman','decimal'} <= page_formats
    settings = F.parse(parts['word/settings.xml'])
    parity = settings.find(F.q('evenAndOddHeaders')) is not None and any(
        s.findall('w:headerReference', F.NS) and {x.get(F.q('type')) for x in s.findall('w:headerReference',F.NS)} >= {'default','even'} for s in sections)
    data_tables = [tid for tid,v in tables.items() if v.get('kind') in ('three-line','complex','panel')]
    panel_tables = [tid for tid,v in tables.items() if v.get('kind') == 'panel']
    table_geometry = all(int(val(ts[tid].find('w:tblPr/w:tblW',F.NS),'w') or '0') > 0 for tid in data_tables)
    continuation_proof = dual and all(dual_evidence.get(e, {}).get('table_spans_verified') is True
                                     and dual_evidence.get(e, {}).get('continued_tables_checked') is True for e in ('word','wps'))
    caption_ps = byrole.get('table_caption',[]) + byrole.get('figure_caption',[])
    equations = byrole.get('equation',[])
    math_count = len(root.findall('.//m:oMath', F.NS))
    legacy_math = [x for x in root.iter() if isinstance(x.tag, str) and x.tag.endswith('}OLEObject') and re.search(r'Equation|MathType|DSMT', x.get('ProgID',''), re.I)]
    eq_layout = all(p.getparent() is body and p.find('w:pPr/w:pBdr',F.NS) is None
                    and {'center','right'} <= {x.get(F.q('val')) for x in p.findall('w:pPr/w:tabs/w:tab',F.NS)}
                    and re.search(r'公式[（(]\d+-\d+[）)]',texts(p)) for p in equations)
    heading_order = [(list(body).index(p), role) for role in ('reference_heading','appendix_heading','acknowledgement_heading')
                     for p in byrole.get(role,[]) if p.getparent() is body]
    ack = [n for n,r in heading_order if r=='acknowledgement_heading']
    back_order = bool(ack and byrole.get('reference_heading')) and all(n<ack[-1] for n,r in heading_order if r!='acknowledgement_heading')
    last_heading = max([list(body).index(p) for role,paragraphs in byrole.items() for p in paragraphs
                        if p.getparent() is body and (role in ('chapter','section','subsection','subsubsection') or role.endswith('_heading'))] or [-1])
    back_order = back_order and ack[-1] == last_heading
    ordinary_empty = []
    code_nodes = set(byrole.get('appendix_code', []))
    first_text_index = min([list(body).index(p) for role in ('cn_abstract_heading','chapter') for p in byrole.get(role,[]) if p.getparent() is body] or [len(body)])
    for i,p in enumerate(body):
        if i < first_text_index or p in code_nodes or p.tag != F.q('p') or F.ptext(p).strip() or p.find('w:pPr/w:sectPr',F.NS) is not None:continue
        if all(x.tag in {F.q(n) for n in ('p','pPr','r','rPr','t')} or any(a.tag in (F.q('pPr'),F.q('rPr')) for a in x.iterancestors()) for x in p.iter()):ordinary_empty.append(i)
    native_math = dual and all(dual_evidence.get(e,{}).get('native_math_count',0) >= math_count for e in ('word','wps'))
    conditions = {
      'page.geometry': (geometry and clean('page.'), True, ['section_geometry']),
      'front.components': (all(has(r) for r in front_roles) and native_cover_logo and not front_geometry_issues and clean('front.','cover.','commitment.'), True, ['cover_roles','commitment_roles','exact_template_logo_in_cover','native_front_template_geometry']),
      'text.styles': (style_bindings and clean('paragraph.','style.'), True, ['effective_styles','final_role_map']),
      'headings.numbering': (heading_native and clean('heading.') and dual, True, ['native_multilevel_numbering','dual_numbering_page_review']),
      'toc.two_levels': (toc_levels and toc_targets and dual and clean('toc.','fields.'), True, ['toc_field_targets','both_engine_cached_target_pages']),
      'section.pagination': (numbering_sections and dual and clean('section.','page_number.','pagination.'), True, ['section_page_number_formats','dual_page_number_review']),
      'section.headers': (parity and dual and clean('header.','header_footer.'), True, ['effective_header_references','displayed_parity_review']),
      'table.structure': (table_geometry and clean('table.') and continuation_proof, bool(data_tables), ['classified_table_geometry','table_audit','native_spans_and_observed_continuation_labels']),
      'table.panel': (clean('table.') and continuation_proof and all(tables[t].get('panel_rows') for t in panel_tables), bool(panel_tables), ['panel_rows_and_dividers','native_spans_and_observed_continuation_labels']),
      'caption.order': (clean('caption.','table.caption','figure.caption') and dual and bool(caption_ps), bool(data_tables or caption_ps), ['caption_roles_sequences','dual_caption_page_review']),
      'math.layout': (eq_layout and clean('math.') and dual, bool(equations), ['native_equation_tabs_labels','dual_formula_page_review']),
      'math.editability': (protection_passed and native_math and not legacy_math, bool(math_count or legacy_math), ['unchanged_expression_structure','both_native_math_collections','dual_formula_page_review']),
      'citations.display': (clean('citation.') and (bool(re.search(r'\[\d+(?:[,，-]\d+)*\]', fulltext))) and dual, has('reference'), ['source_bound_citation_audit','complete_marker_and_baseline_review']),
      'references.layout': (has('reference') and has('reference_heading') and clean('reference.') and protection_passed and dual, True, ['bibliography_roles','source_bound_entry_protection','dual_reference_page_review']),
      'backmatter.order': (back_order and clean('backmatter.','structure.') and protection_passed, True, ['last_heading_is_acknowledgement','whole_block_protection']),
      'whitespace.layout': (not ordinary_empty and protection_passed and dual, True, ['no_removable_body_spacer_paragraphs','recorded_exact_edits','dual_spacing_review']),
      'integrity.objects': (protection_passed, True, ['fresh_registered_stage_replay','independent_research_and_object_checks']),
      'delivery.dual_engine': (dual, True, ['same_final_hash','word_all_pages','wps_all_pages']),
    }
    observed = {
      'page.geometry': {'sections': len(sections), 'all_match_spec': geometry},
      'front.components': {'present_roles': [r for r in front_roles if has(r)], 'exact_template_logo_in_cover': native_cover_logo,'template_geometry_issues':front_geometry_issues},
      'text.styles': {'mapped_paragraphs': len(roles), 'all_styles_resolve': style_bindings},
      'headings.numbering': {'heading_paragraphs': len(heading_ps), 'native_numbering': heading_native},
      'toc.two_levels': {'instruction': toc_code, 'required_targets_present': toc_targets, 'levels_at_most_two': toc_levels},
      'section.pagination': {'section_count': len(sections), 'number_formats': sorted(x for x in page_formats if x)},
      'section.headers': {'odd_even_references_present': parity},
      'table.structure': {'table_ids': data_tables, 'reviewed_settings': {t:tables[t] for t in data_tables}, 'continuation_proof': continuation_proof},
      'table.panel': {'table_ids': panel_tables, 'continuation_proof': continuation_proof},
      'caption.order': {'caption_paragraphs': len(caption_ps), 'data_tables': len(data_tables)},
      'math.layout': {'equation_paragraphs': len(equations), 'native_layout_valid': eq_layout},
      'math.editability': {'omml_count': math_count, 'legacy_object_count': len(legacy_math), 'native_math_collections_cover_all': native_math},
      'citations.display': {'numeric_markers_present': bool(re.search(r'\[\d+(?:[,，-]\d+)*\]',fulltext))},
      'references.layout': {'entries': len(byrole.get('reference',[])), 'heading_present': has('reference_heading')},
      'backmatter.order': {'block_heading_order': heading_order, 'acknowledgement_is_last': back_order},
      'whitespace.layout': {'removable_body_spacer_indices': ordinary_empty},
      'integrity.objects': {'fresh_chain_and_independent_protection_passed': protection_passed},
      'delivery.dual_engine': {'same_final_sha256': final_hash, 'dual_all_pages_verified': dual},
    }
    results=[]
    for rule in registry['rules']:
        if rule['id'] not in conditions:
            results.append({'rule_id':rule['id'],'status':'fail','reason':'No implemented evaluator for an enabled rule','evidence':[]})
            continue
        passed,applicable,evidence=conditions[rule['id']]
        results.append({'rule_id':rule['id'],'status':'not_applicable' if not applicable else 'pass' if passed else 'fail',
                        'evidence':evidence, 'observed':observed[rule['id']], 'audit_bound_to_final':audit_bound,
                        'reason':None if applicable else 'No matching object exists in this exact final document'})
    return {'profile_id':registry['profile_id'],'rule_results':results,'passed':all(r['status']!='fail' for r in results),
            'content_advice_only':registry.get('content_advice_only',[])}
