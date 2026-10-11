#!/usr/bin/env python3
"""Read-only structural checks for a newly authored FIT thesis example.

This does not replace native Word field updates, PDF rendering or visual QA.
The caller must verify actual ListString values and physical/logical parity.
"""
from __future__ import annotations

import argparse
import json
import re
import zipfile
from pathlib import Path
from lxml import etree as E

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
N = {"w": W, "m": M, "r": R}


def check(path):
    with zipfile.ZipFile(path) as package:
        files = {name: package.read(name) for name in package.namelist()}
    parser = E.XMLParser(resolve_entities=False, no_network=True)
    document = E.fromstring(files["word/document.xml"], parser)
    styles = E.fromstring(files["word/styles.xml"], parser)
    numbering = E.fromstring(files["word/numbering.xml"], parser)
    settings = E.fromstring(files["word/settings.xml"], parser)
    errors = []
    headings = []
    # Native Word SaveAs is allowed to rename internal style IDs. Resolve by
    # the public names rather than reporting a functional document as broken.
    names = {style.xpath("string(w:name/@w:val)", namespaces=N): style.get("{" + W + "}styleId") for style in styles.xpath("w:style", namespaces=N)}
    heading_names = ("论文-章标题", "论文-一级节标题", "论文-二级节标题", "论文-三级节标题")
    ids = tuple(names.get(name, "MISSING_" + str(i)) for i, name in enumerate(heading_names))
    math_id = names.get("论文-数学变量", "MISSING_MATH")
    label_id = names.get("论文-公式编号文字", "MISSING_LABEL")
    equation_id = names.get("论文-公式", "MISSING_EQUATION")
    style_index = {style.get("{" + W + "}styleId"): style for style in styles.xpath("w:style", namespaces=N)}
    def effective_number(paragraph, style_id):
        values = {}
        sources = [paragraph]
        visited = set()
        while style_id in style_index and style_id not in visited:
            visited.add(style_id)
            style = style_index[style_id]
            sources.append(style)
            style_id = style.xpath("string(w:basedOn/@w:val)", namespaces=N)
        for source in sources:
            for key in ("ilvl", "numId"):
                found = source.xpath("w:pPr/w:numPr/w:" + key + "/@w:val", namespaces=N)
                if found and key not in values:
                    values[key] = found[0]
        return values.get("ilvl", "0"), values.get("numId", "")
    for paragraph in document.xpath(".//w:p", namespaces=N):
        sid = paragraph.xpath("string(w:pPr/w:pStyle/@w:val)", namespaces=N)
        if sid not in ids:
            continue
        text = "".join(paragraph.xpath(".//w:t/text()", namespaces=N))
        if re.match(r"^(?:[一二三四五六七八九十百]+、|（[一二三四五六七八九十百0-9]+）|\d+[.、])", text):
            errors.append("Manual heading number: " + text)
        level, num_id = effective_number(paragraph, sid)
        if level != str(ids.index(sid)) or not num_id or num_id == "0":
            errors.append("Missing/incorrect live heading numPr: " + text)
        num = numbering.xpath(f'w:num[@w:numId="{num_id}"]/w:abstractNumId/@w:val', namespaces=N)
        if len(num) != 1:
            errors.append("Heading refers to absent numbering definition: " + text)
        else:
            actual = numbering.xpath(f'w:abstractNum[@w:abstractNumId="{num[0]}"]/w:lvl[@w:ilvl="{level}"]/w:lvlText/@w:val', namespaces=N)
            expected = ("%1、", "（%2）", "%3.", "（%4）")[ids.index(sid)]
            if actual != [expected]:
                errors.append("Heading list does not use school display format: " + text)
        if sid == ids[0]:
            fields = paragraph.xpath(".//w:instrText/text()", namespaces=N)
            if not any("SEQ FITChapter" in f and "\\h" in f for f in fields):
                errors.append("Arabic chapter sequence marker absent: " + text)
            if not any("SEQ FITEquation" in f and "\\r 0" in f for f in fields):
                errors.append("Equation chapter reset marker absent: " + text)
        headings.append({"text": text, "style": sid, "level": level, "num_id": num_id})
    if not settings.xpath("w:evenAndOddHeaders[not(@w:val='0') and not(@w:val='false')]", namespaces=N):
        errors.append("Real odd/even header setting absent")
    sections = []
    for index, section in enumerate(document.xpath(".//w:sectPr", namespaces=N), 1):
        headers = section.xpath("w:headerReference/@w:type", namespaces=N)
        footers = section.xpath("w:footerReference/@w:type", namespaces=N)
        if not {"default", "even"}.issubset(headers):
            errors.append(f"Section {index} lacks actual default/even headers")
        # The original school cover/commitment region intentionally has no
        # footer references at all. Absence in this first unpaged region is
        # valid; paginated regions must have both stories.
        unpaged_first = index == 1 and not section.xpath("w:pgNumType", namespaces=N)
        if not unpaged_first and not {"default", "even"}.issubset(footers):
            errors.append(f"Section {index} lacks actual default/even footers")
        if section.xpath("w:titlePg[not(@w:val='0') and not(@w:val='false')]", namespaces=N):
            errors.append(f"Section {index} has different-first-page enabled")
        sections.append({"index": index, "headers": headers, "footers": footers, "format": section.xpath("string(w:pgNumType/@w:fmt)", namespaces=N), "start": section.xpath("string(w:pgNumType/@w:start)", namespaces=N)})
    instructions = document.xpath(".//w:instrText/text()", namespaces=N)
    equations = [s for s in instructions if re.match(r"^\s*SEQ\s+FITEquation(?:\s|$)", s) and "\\h" not in s and "\\r" not in s]
    formula_paragraphs = []
    for paragraph in document.xpath(".//w:p", namespaces=N):
        sequence_fields = [s for s in paragraph.xpath(".//w:instrText/text()", namespaces=N) if re.match(r"^\s*SEQ\s+FITEquation(?:\s|$)", s) and "\\h" not in s and "\\r" not in s]
        if sequence_fields:
            formula_paragraphs.append(paragraph)
    if len(equations) != len(formula_paragraphs):
        errors.append("Equation label count differs from live equation sequence count")
    def effective_property(paragraph, property_name):
        found = paragraph.find("{" + W + "}pPr/{" + W + "}" + property_name)
        if found is not None:
            return found
        sid = paragraph.xpath("string(w:pPr/w:pStyle/@w:val)", namespaces=N)
        visited = set()
        while sid in style_index and sid not in visited:
            visited.add(sid)
            found = style_index[sid].find("{" + W + "}pPr/{" + W + "}" + property_name)
            if found is not None:
                return found
            sid = style_index[sid].xpath("string(w:basedOn/@w:val)", namespaces=N)
        return None
    for paragraph in formula_paragraphs:
        if paragraph.xpath("ancestor::w:tbl|ancestor::w:txbxContent|ancestor::w:drawing|ancestor::w:pict", namespaces=N):
            errors.append("Equation is inside a table, frame or text box; use a plain tabbed paragraph")
        if effective_property(paragraph, "pBdr") is not None or effective_property(paragraph, "framePr") is not None:
            errors.append("Equation paragraph has borders or frame properties")
        if paragraph.xpath("string(w:pPr/w:pStyle/@w:val)", namespaces=N) != equation_id:
            errors.append("Equation does not use the named formula paragraph style")
        if not paragraph.xpath("m:oMath", namespaces=N):
            errors.append("Equation lacks native OMML in its plain paragraph")
        if len(paragraph.xpath("w:r/w:tab", namespaces=N)) != 2:
            errors.append("Equation needs a leading center tab and a label right tab")
        tabs = effective_property(paragraph, "tabs")
        tab_values = [(x.get("{" + W + "}val"), int(x.get("{" + W + "}pos", "0"))) for x in tabs] if tabs is not None else []
        center = [p for v, p in tab_values if v == "center"]
        right = [p for v, p in tab_values if v == "right"]
        if len(center) != 1 or len(right) != 1 or abs(center[0] * 2 - right[0]) > 1:
            errors.append("Equation center/right tab stops do not share the text-width midpoint")
        alignment = effective_property(paragraph, "jc")
        if alignment is not None and alignment.get("{" + W + "}val") != "left":
            errors.append("Tab-aligned equation paragraph must use left paragraph alignment")
        text = "".join(paragraph.xpath(".//w:t/text()", namespaces=N))
        if not re.fullmatch(r"公式（\d+-\d+）", text):
            errors.append("Incomplete equation label cache: " + text)
        label_runs = [r for r in paragraph.xpath("w:r", namespaces=N) if r.xpath("w:t|w:instrText|w:fldChar", namespaces=N)]
        if any(run.xpath("string(w:rPr/w:rStyle/@w:val)", namespaces=N) != label_id for run in label_runs):
            errors.append("Equation label runs lack the editable label character style")
    for run in document.xpath(".//m:r", namespaces=N):
        if run.xpath("string(w:rPr/w:rStyle/@w:val)", namespaces=N) != math_id:
            errors.append("Math run lacks editable mathematics character style")
            break
        if run.xpath("w:rPr/w:rFonts|w:rPr/w:sz|w:rPr/w:szCs", namespaces=N):
            errors.append("Math run has direct font/size; style propagation requires explicit review")
            break
    for sid in ids + (math_id, label_id, equation_id):
        found = styles.xpath(f'w:style[@w:styleId="{sid}"]', namespaces=N)
        if not found or not found[0].xpath("w:qFormat", namespaces=N):
            errors.append("Named style absent/not visible: " + sid)
    for name, data in files.items():
        if re.fullmatch(r"word/header\d+\.xml", name):
            header = E.fromstring(data, parser)
            if any(re.search(r"\bIF\b", f) for f in header.xpath(".//w:instrText/text()", namespaces=N)):
                errors.append("Conditional IF cannot replace genuine odd/even headers: " + name)
    return {"path": str(Path(path).resolve()), "passed": not errors, "errors": errors, "headings": headings, "sections": sections, "equation_labels": len(equations), "equation_layout": "plain-tabbed-paragraph", "formula_tables": sum(bool(p.xpath("ancestor::w:tbl", namespaces=N)) for p in formula_paragraphs), "native_math": len(document.xpath(".//m:oMath", namespaces=N)), "math_scripts": {name: len(document.xpath(".//m:" + name, namespaces=N)) for name in ("f", "sSub", "sSup", "sSubSup", "nary", "d", "rad")}, "note": "Requires native Word field/list/page-parity checks and rendered-page visual inspection."}


if __name__ == "__main__":
    arguments = argparse.ArgumentParser(description=__doc__)
    arguments.add_argument("docx")
    arguments.add_argument("--report")
    args = arguments.parse_args()
    result = check(args.docx)
    if args.report:
        Path(args.report).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    raise SystemExit(0 if result["passed"] else 1)
