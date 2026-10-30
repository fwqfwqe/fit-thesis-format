#!/usr/bin/env python3
"""Read-only, evidence-based OOXML format audit for FIT economics theses.

This script never opens Word, updates fields, reconstructs content or writes the
input package. XML IDs locate evidence; they are not rendered page numbers.
"""
from __future__ import annotations

import argparse
from collections import Counter, defaultdict
from datetime import datetime, timezone
import hashlib
import json
import posixpath
from pathlib import Path
import re
import sys
import zipfile
from lxml import etree as E
from template_hygiene import inspect_parts as inspect_review_parts

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"w": W, "m": M, "r": R, "a": A,
      "o": "urn:schemas-microsoft-com:office:office"}
SCHOOL_HEADER = "福州理工学院本科生毕业论文（设计）"
NUMBER_TARGETS = {
    "chapter": (0, "chineseCounting", "%1、"),
    "section": (1, "chineseCounting", "（%2）"),
    "subsection": (2, "decimal", "%3."),
    "subsubsection": (3, "decimal", "（%4）"),
}
ROLE_ALIASES = {"body_math": "body", "table_math": "table_text",
                "equation_number": "equation"}
PARA_CONTROLS = {"jc", "spacing", "ind", "outlineLvl", "pageBreakBefore",
                 "keepNext", "keepLines"}
RUN_CONTROLS = {"rFonts", "sz", "szCs"}
FORMAT_STATUSES = {
    "auto_fixable": "确认角色后可自动修复格式",
    "requires_content_authorization": "涉及字符/域显示，需核对现有授权及转换记录",
    "requires_render_review": "须在 Word/PDF 渲染后人工复核",
    "requires_role_review": "须先确认语义角色或数据表类型",
}


def q(n): return "{" + W + "}" + n
def local(n): return E.QName(n).localname
def sha(b): return hashlib.sha256(b).hexdigest()
def parse(b):
    return E.fromstring(b, E.XMLParser(resolve_entities=False, no_network=True))
def val(e, name="val", default=None):
    return e.get(q(name), default) if e is not None else default
def enabled(e): return e is not None and val(e, default="1") not in {"0", "false", "off"}
def text(e): return "".join(e.xpath(".//w:t/text()", namespaces=NS))
def cached_visible_text(e):
    """Exclude nested field results embedded in an outer field instruction."""
    stack, chunks = [], []
    for el in e.iter():
        if el.tag == q("fldChar"):
            kind = val(el, "fldCharType")
            if kind == "begin": stack.append("instruction")
            elif kind == "separate" and stack: stack[-1] = "result"
            elif kind == "end" and stack: stack.pop()
        elif el.tag == q("t") and all(state == "result" for state in stack):
            chunks.append(el.text or "")
    return "".join(chunks)
def clean(s): return re.sub(r"\s+", "", s or "")
def short(s, length=90): return s[:length] + ("…" if len(s) > length else "")
def number(v):
    try: return float(v)
    except (TypeError, ValueError): return None
def props(e):
    if e is None: return {}
    result = {}
    for c in e:
        if isinstance(c.tag, str):
            result[local(c)] = {local(k): v for k, v in c.attrib.items()} or {"val": "1"}
    return result
def merge(*items):
    result = {}
    for item in items:
        for k, v in item.items():
            current = result.setdefault(k, {})
            if k == "rFonts":
                # A concrete font in a later cascade layer overrides an inherited
                # theme font. Theme+literal in the same layer retains theme's
                # OOXML precedence; retaining an earlier theme caused false
                # Calibri reports on native Word-saved Times New Roman styles.
                for name in ("ascii", "hAnsi", "eastAsia", "cs"):
                    if name in v and name+"Theme" not in v: current.pop(name+"Theme", None)
            current.update(v)
    return result
def pvalue(pr, name, attr="val", default=None): return pr.get(name, {}).get(attr, default)
def fields(container):
    """Extract each field's own instruction, including nested field children."""
    result, stack = [], []
    for el in container.iter():
        if el.tag == q("fldSimple"):
            result.append({"instruction": val(el, "instr", ""), "simple": True,
                           "complete": True, "cached_text": text(el), "nested": []})
        elif el.tag == q("fldChar"):
            kind = val(el, "fldCharType")
            if kind == "begin":
                record = {"instruction": "", "simple": False, "complete": False,
                          "cached_text": "", "nested": [], "in_result": False}
                if stack: stack[-1]["nested"].append(record)
                stack.append(record); result.append(record)
            elif kind == "separate" and stack: stack[-1]["in_result"] = True
            elif kind == "end" and stack: stack.pop()["complete"] = True
        elif el.tag == q("instrText") and stack:
            stack[-1]["instruction"] += el.text or ""
        elif el.tag == q("t") and stack and stack[-1]["in_result"]:
            stack[-1]["cached_text"] += el.text or ""
    for f in result: f.pop("in_result", None)
    return result


def plain_math_candidates(p):
    """Return review candidates, never an instruction to transcribe mathematics.

    Only ordinary text is considered. Paths, URLs and code-looking runs are
    excluded; an underscore alone is not enough to identify a formula.
    """
    chunks = []
    for run in p.iter(q("r")):
        if any(a.tag.startswith("{" + M + "}") for a in run.iterancestors()):
            continue
        value = text(run)
        if value:
            chunks.append(value)
    value = "".join(chunks)
    value = re.sub(r"(?:https?://|www\.)\S+|[A-Za-z]:[\\/]\S+|\S+[\\/]\S+", " ", value)
    value = re.sub(r"\b[A-Za-z0-9_]+\.(?:py|csv|json|txt|docx|pdf|xlsx)\b", " ", value, flags=re.I)
    if re.search(r"^\s*(?:def |import |from \S+ import |for |if |return |SELECT |UPDATE )", value, re.I):
        return []
    atom = r"[A-Za-zΑ-Ωα-ω][A-Za-zΑ-Ωα-ω0-9]*(?:[_^](?:\{[A-Za-z0-9,+ -]+\}|[A-Za-z0-9]+))*"
    relation = re.compile(atom + r"\s*[=＝≈≠≤≥]\s*(?:" + atom + r"|[-−]?\d+(?:[.．]\d+)?(?:[eE][+-]?\d+)?)")
    candidates = [m.group() for m in relation.finditer(value)]
    if re.search(r"变量|系数|公式|函数|模型|下标|上标|\b(?:variable|coefficient|equation|model)\b", value, re.I):
        for match in re.finditer(atom, value):
            token = match.group()
            if "_" not in token and "^" not in token: continue
            base, suffix = re.split(r"[_^]", token, maxsplit=1)
            conventional = len(base) <= 3 and re.fullmatch(r"\{?[a-z0-9,+ -]{1,6}\}?", suffix)
            adjacent_definition = re.search(r"(?:变量|系数|下标|上标|variable|coefficient)\s*(?:为|是|如|包括)?\s*$", value[max(0,match.start()-20):match.start()], re.I)
            if conventional or adjacent_definition:
                candidates.append(token)
    return list(dict.fromkeys(candidates))[:8]


class Audit:
    def __init__(self, source, mapping=None, title=None):
        self.source = Path(source).resolve()
        self.before = self.source.read_bytes()
        self.hash = sha(self.before)
        with zipfile.ZipFile(self.source) as z:
            if len(z.namelist()) != len(set(z.namelist())):
                raise ValueError("DOCX 包含重名 ZIP 部件，无法可靠审核。")
            self.parts = {n: z.read(n) for n in z.namelist()}
        if "word/document.xml" not in self.parts or "word/styles.xml" not in self.parts:
            raise ValueError("不是可审核的 Word DOCX 文档。")
        self.doc = parse(self.parts["word/document.xml"])
        self.styles_root = parse(self.parts["word/styles.xml"])
        self.styles = {val(s, "styleId"): s for s in self.styles_root.findall(q("style"))}
        self.default_style = next((sid for sid, s in self.styles.items()
                                   if val(s, "type") == "paragraph" and enabled_attr(s, "default")), "Normal")
        self.default_p = props(self.styles_root.find("w:docDefaults/w:pPrDefault/w:pPr", NS))
        self.default_r = props(self.styles_root.find("w:docDefaults/w:rPrDefault/w:rPr", NS))
        self.numbering = parse(self.parts["word/numbering.xml"]) if "word/numbering.xml" in self.parts else None
        self.settings = parse(self.parts["word/settings.xml"]) if "word/settings.xml" in self.parts else E.Element(q("settings"))
        self.theme = parse(self.parts["word/theme/theme1.xml"]) if "word/theme/theme1.xml" in self.parts else None
        self.spec_path = Path(__file__).resolve().parent.parent / "references" / "style-spec.json"
        self.spec_bytes = self.spec_path.read_bytes()
        self.spec = json.loads(self.spec_bytes.decode("utf-8-sig"))
        self.ps = {f"p{i:06d}": p for i, p in enumerate(self.doc.iter(q("p")), 1)}
        self.ts = {f"t{i:04d}": t for i, t in enumerate(self.doc.iter(q("tbl")), 1)}
        self.pid = {p: pid for pid, p in self.ps.items()}
        self.tid = {t: tid for tid, t in self.ts.items()}
        self.mapping = mapping or {}
        if self.mapping and self.mapping.get("source_sha256") != self.hash:
            raise ValueError("role-map 的 source_sha256 不匹配；不能在另一版本论文上沿用定位。")
        for pid in self.mapping.get("paragraphs", {}):
            if pid not in self.ps: raise ValueError("role-map 有不存在的段落：" + pid)
        for tid in self.mapping.get("tables", {}):
            if tid not in self.ts: raise ValueError("role-map 有不存在的表格：" + tid)
        self.issues, self.roles = [], {}
        self.evidence = defaultdict(list)
        self.style_cache = {}
        self.title_override = title
        self.title = None
        self.front_end = 0
        self.body_start = None
        self.sections = []
        self.checked_story_formats = set()
        self.toc_pids = set()
        stack = []
        for el in self.doc.iter():
            if el.tag == q("fldSimple") and re.match(r"\s*TOC\b", val(el, "instr", ""), re.I):
                self.toc_pids.update(self.pid[p] for p in el.iter(q("p")) if p in self.pid)
            elif el.tag == q("fldChar"):
                kind = val(el, "fldCharType")
                if kind == "begin": stack.append({"instruction": "", "toc": False})
                elif kind == "end" and stack: stack.pop()
            elif el.tag == q("instrText") and stack:
                stack[-1]["instruction"] += el.text or ""
                stack[-1]["toc"] = bool(re.match(r"\s*TOC\b", stack[-1]["instruction"], re.I))
            elif el.tag == q("t") and any(f["toc"] for f in stack):
                p = next((a for a in el.iterancestors() if a.tag == q("p")), None)
                if p in self.pid: self.toc_pids.add(self.pid[p])

    def add(self, severity, category, location, actual, expected, recommendation,
            status="auto_fixable", confidence="high"):
        self.issues.append({"id": f"F{len(self.issues)+1:05d}", "severity": severity,
                            "category": category, "location": location,
                            "actual": actual, "expected": expected,
                            "recommendation": recommendation, "status": status,
                            "confidence": confidence})

    def locate(self, pid=None, tid=None, sid=None, part="word/document.xml", **more):
        loc = {"part": part}
        if pid:
            p = self.ps.get(pid)
            loc.update({"paragraph_id": pid, "xml_path": f"(//w:p)[{int(pid[1:])}]"})
            if p is not None:
                loc["text_excerpt"] = short(text(p))
                table = next((x for x in p.iterancestors() if x.tag == q("tbl")), None)
                if table is not None: loc["table_id"] = self.tid.get(table)
                if pid in self.roles:
                    loc.update({"role": self.roles[pid][0], "role_basis": self.roles[pid][2]})
        if tid: loc.update({"table_id": tid, "xml_path": f"(//w:tbl)[{int(tid[1:])}]"})
        if sid: loc["section_id"] = sid
        loc.update(more)
        return loc

    def style_properties(self, sid):
        if sid in self.style_cache: return self.style_cache[sid]
        chain, seen, current = [], set(), sid
        while current and current in self.styles and current not in seen:
            seen.add(current); st = self.styles[current]; chain.append(st)
            current = val(st.find(q("basedOn")))
        # A character style inherits only explicitly defined character-style
        # properties. Applying docDefaults a second time would incorrectly hide
        # the paragraph font underneath a superscript character style.
        character = sid in self.styles and val(self.styles[sid], "type") == "character"
        pp = merge({} if character else self.default_p, *[props(s.find(q("pPr"))) for s in reversed(chain)])
        rp = merge({} if character else self.default_r, *[props(s.find(q("rPr"))) for s in reversed(chain)])
        self.style_cache[sid] = (pp, rp)
        return pp, rp

    def sid(self, p): return val(p.find("w:pPr/w:pStyle", NS), default=self.default_style)
    def effective_p(self, p): return merge(self.style_properties(self.sid(p))[0], props(p.find(q("pPr"))))
    def effective_paragraph_container(self, p, name):
        result, chain, seen, current = None, [], set(), self.sid(p)
        while current in self.styles and current not in seen:
            seen.add(current); st = self.styles[current]; chain.append(st)
            current = val(st.find(q("basedOn")))
        for st in reversed(chain):
            item = st.find("w:pPr/w:"+name, NS)
            if item is not None: result = item
        item = p.find("w:pPr/w:"+name, NS)
        return item if item is not None else result

    def equation_tab_layout(self, p):
        tabs = self.effective_paragraph_container(p, "tabs")
        if tabs is None: return False
        centers = [number(val(t, "pos")) for t in tabs.findall(q("tab")) if val(t) == "center"]
        rights = [number(val(t, "pos")) for t in tabs.findall(q("tab")) if val(t) == "right"]
        width = self.spec["page"]["width_twips"] - self.spec["page"]["left_twips"] - self.spec["page"]["right_twips"]
        has_center = any(v is not None and abs(v-width/2) <= 40 for v in centers)
        has_right = any(v is not None and abs(v-width) <= 40 for v in rights)
        return has_center and has_right and len(p.findall(".//w:tab", NS)) >= 2
    def effective_r(self, p, r):
        rp = r.find(q("rPr")); cs = val(rp.find(q("rStyle"))) if rp is not None else None
        cp = self.style_properties(cs)[1] if cs else {}
        return merge(self.style_properties(self.sid(p))[1], cp, props(rp))

    def font(self, rp, category):
        f = rp.get("rFonts", {})
        key = "eastAsia" if category == "east_asia" else "ascii"
        token = f.get(key + "Theme")
        if token and self.theme is not None:
            group = "majorFont" if token.lower().startswith("major") else "minorFont"
            fs = self.theme.find("a:themeElements/a:fontScheme/a:" + group, NS)
            if fs is not None:
                ft = fs.find("a:ea" if category == "east_asia" else "a:latin", NS)
                if ft is not None and ft.get("typeface"): return ft.get("typeface")
                if category == "east_asia":
                    ft = fs.find("a:font[@script='Hans']", NS)
                    if ft is not None: return ft.get("typeface")
            return "主题字体未能解析:" + token
        return f.get(key) or (f.get("hAnsi") if category != "east_asia" else None)

    def infer_roles(self):
        pmap = self.mapping.get("paragraphs", {})
        names_to_roles = {v["name"]: k for k, v in self.spec["roles"].items()}
        names_to_roles.update({"论文-正文含公式": "body_math", "论文-表内含公式": "table_math",
                               "论文-公式编号": "equation_number", "论文-三级目录": "toc3",
                               "论文-附录代码": "appendix_code"})
        toc_region, abstract_region, ref_region, appendix = False, None, False, False
        cn_abstract_idx = None
        for index, (pid, p) in enumerate(self.ps.items(), 1):
            s, sid = text(p).strip(), self.sid(p)
            style = self.styles.get(sid)
            name = val(style.find(q("name"))) if style is not None else ""
            role, confidence, basis = None, "medium", "文本和结构启发式"
            item = pmap.get(pid)
            if item is not None:
                role = item if isinstance(item, str) else item["role"]
                if isinstance(item, dict) and "expected_text" in item and item["expected_text"] != text(p):
                    raise ValueError("role-map 的 expected_text 不匹配：" + pid)
                confidence, basis = "high", "显式、SHA-256 绑定角色映射"
            elif sid.startswith("FITRepair_"):
                from format_engine import _read_repair_style
                repaired = _read_repair_style(style)
                role, confidence, basis = repaired[0], "high", "来源绑定的分页修复样式"
            elif sid.startswith("FITCode_"):
                role, confidence, basis = "appendix_code", "high", "保留原字体及缩进的附录代码样式"
            elif sid.startswith("FIT_"):
                role, confidence, basis = sid[4:], "high", "已有 FIT 命名样式"
            elif name in names_to_roles:
                role, confidence, basis = names_to_roles[name], "high", "已有学校语义命名样式（Word 可改写 styleId）"
            elif re.match(r"^(?:TOC|目录)\s*[123]$", name or "", re.I) or pid in self.toc_pids:
                match_level = re.search(r"[123]$", name or "")
                role = "toc" + (match_level.group() if match_level else "1")
                confidence, basis = "high", "真实 TOC 域缓存区域或 TOC 目录样式"
            elif clean(s) in {"摘要", "中文摘要"}: role = "cn_abstract_heading"
            elif clean(s).lower() == "abstract": role = "en_abstract_heading"
            elif clean(s) in {"目录", "目次"}: role = "toc_heading"
            elif clean(s) == "参考文献": role = "reference_heading"
            elif clean(s) == "致谢": role = "acknowledgement_heading"
            elif clean(s) in {"结论", "研究结论", "结论与展望"}: role = "conclusion_heading"
            elif re.fullmatch(r"附录(?:[一二三四五六七八九十A-Z0-9]+)?", clean(s)): role = "appendix_heading"
            elif re.match(r"^(?:关键词|关键字)\s*[:：]", s): role = "cn_keywords"
            elif re.match(r"^key\s*words?\s*[:：]", s, re.I): role = "en_keywords"
            elif re.match(r"^(?:数据来源|资料来源)\s*[:：]", s): role = "source"
            elif re.match(r"^表\s*\d+\s*[-－—]\s*\d+", s): role = "table_caption"
            elif re.match(r"^图\s*\d+\s*[-－—]\s*\d+", s): role = "figure_caption"
            elif s and (re.match(r"^(?:heading|标题)\s*[1-4]$", name or "", re.I)
                        or pvalue(self.effective_p(p), "outlineLvl") in {"0", "1", "2", "3"}):
                level = pvalue(self.effective_p(p), "outlineLvl")
                if level not in {"0", "1", "2", "3"}:
                    level = str(int(re.search(r"[1-4]$", name).group()) - 1)
                role, confidence, basis = list(NUMBER_TARGETS)[int(level)], "high", "标题样式或大纲级别"
            elif s and re.match(r"^[一二三四五六七八九十]+、\S", s) and len(s) <= 55:
                role, confidence = "chapter", "medium"
            elif s and re.match(r"^（[一二三四五六七八九十]+）\S", s) and len(s) <= 55:
                role, confidence = "section", "medium"
            elif p.findall(".//m:oMathPara", NS) and not s:
                role = "equation"
            elif p.findall(".//w:drawing", NS) and not s:
                role = "figure"
            if item is None and appendix and role in NUMBER_TARGETS and re.fullmatch(r"附录[一二三四五六七八九十A-Z0-9]+", clean(s)):
                # A label inside an existing appendix/code container is not a
                # main-body section just because it was assigned that style.
                role, confidence, basis = "appendix_subheading", "high", "已有附录内的分项标题"
            if item is None and role in NUMBER_TARGETS and len(clean(s)) > 120:
                # A heading style is evidence, not proof that a long prose
                # paragraph is a heading. Keep the candidate for review;
                # never silently turn it into body text or rewrite it.
                confidence = "medium"
                self.add("warning", "heading.long_text_requires_review", self.locate(pid=pid),
                         {"candidate_role": role, "style_id": sid, "text_length": len(clean(s))},
                         "长段落的标题角色已结合原文确认",
                         "该段落带标题样式或大纲级别，但文字较长。阅读完整原文后在角色映射中明确确认；不要仅凭样式把散文当标题或自动改为正文。",
                         "requires_role_review", "medium")
            if role == "cn_abstract_heading":
                cn_abstract_idx = index; abstract_region = "cn_abstract"; toc_region = ref_region = False
            elif role == "en_abstract_heading":
                abstract_region = "en_abstract"; toc_region = ref_region = False
            elif role == "toc_heading":
                toc_region = True; abstract_region = None; ref_region = False
            elif role == "reference_heading":
                ref_region = True; abstract_region = None; toc_region = False
            elif role in {"acknowledgement_heading", "appendix_heading"}:
                ref_region = False; abstract_region = None; appendix = role == "appendix_heading"
            elif role in NUMBER_TARGETS:
                if not toc_region or confidence == "high":
                    toc_region = False; abstract_region = None
                    appendix = False
                    if self.body_start is None and not appendix: self.body_start = index
            if role is None and s:
                inside_table = any(a.tag == q("tbl") for a in p.iterancestors())
                if ref_region and re.match(r"^\[\s*\d+\s*\]", s): role = "reference"
                elif toc_region:
                    # A PAGEREF makes this an entry; a plain digit alone does not.
                    if any(re.match(r"\s*PAGEREF\b", f["instruction"], re.I) for f in fields(p)):
                        role = "toc1"
                elif abstract_region: role = abstract_region
                elif inside_table: role, confidence = "table_unclassified", "low"
                elif self.body_start and index >= self.body_start and not appendix:
                    role = "body"
            if any(a.tag in {q('drawing'),q('pict'),q('object')} for a in p.iterancestors()):
                role,confidence,basis='protected_object_text','high','嵌入图形内部文字，保留原对象并逐页复核'
            if role:
                self.roles[pid] = (role, confidence, basis)
                self.evidence[role].append(pid)
        if cn_abstract_idx:
            self.front_end = cn_abstract_idx - 1
            # School/competition cover emblems are not scientific figures.
            for pid, (role, confidence, basis) in list(self.roles.items()):
                if int(pid[1:]) < cn_abstract_idx and role == "figure" and basis == "文本和结构启发式":
                    self.roles[pid] = ("cover_image_unclassified", "low", "前置封面图形，未作正文图片处理")
        title_candidates = self.evidence["cn_title"] + self.evidence["cover_title"]
        if self.title_override:
            self.title = self.title_override
        elif title_candidates:
            self.title = text(self.ps[title_candidates[0]]).strip()
        elif cn_abstract_idx:
            prior = list(self.ps.items())[:cn_abstract_idx-1]
            for _, p in reversed(prior):
                s = text(p).strip()
                if s and len(s) > 5 and len(s) < 160 and not re.match(r"^(姓名|题目|学院|专业|学号|年级|指导教师|时间)\s*[:：]?$", s):
                    self.title = s; break

    def check_coverage(self):
        """Make every substantive unclassified paragraph visible to delivery.

        Heuristics deliberately do not guess the roles of appendix prose,
        unknown heading conventions or content after an uncertain TOC. These
        omissions must be review findings, not a silent lack of format errors.
        Empty spacing paragraphs do not require a semantic role.
        """
        content_tags = {q(name) for name in ("t", "delText", "instrText", "delInstrText",
                        "drawing", "pict", "object", "fldChar", "fldSimple", "sym",
                        "footnoteReference", "endnoteReference")} | {"{" + M + "}oMath"}
        text_tags = {q(name) for name in ("t", "delText", "instrText", "delInstrText")}
        substantive = []
        for pid, p in self.ps.items():
            for node in p.iter():
                if node.tag not in content_tags:
                    continue
                if node.tag == q('fldChar') and val(node, 'fldCharType') == 'end':
                    # Word can place a TOC closing marker in its own empty
                    # paragraph. It has no displayed content to classify;
                    # the full review plan still retains this control node.
                    continue
                # Textbox paragraphs have their own IDs; do not count their
                # text again as ordinary prose in the outer paragraph.
                if next((a for a in node.iterancestors() if a.tag == q("p")), None) is not p:
                    continue
                if node.tag not in text_tags or (node.text or "").strip():
                    substantive.append(pid)
                    break
        missing = [pid for pid in substantive if pid not in self.roles]
        ranges = []
        for pid in missing:
            if ranges and int(pid[1:]) == int(ranges[-1]["last"][1:]) + 1:
                ranges[-1]["last"] = pid
                ranges[-1]["count"] += 1
            else:
                ranges.append({"first": pid, "last": pid, "count": 1})
        self.coverage = {"part": "word/document.xml", "substantive_paragraph_count": len(substantive),
                         "classified_substantive_paragraph_count": len(substantive) - len(missing),
                         "unclassified_paragraph_ids": missing, "unclassified_ranges": ranges,
                         "role_review_required": bool(missing)}
        if missing:
            self.add("warning", "coverage.unclassified_content",
                     self.locate(pid=missing[0], paragraph_range=[missing[0], missing[-1]]),
                     {"unclassified_count": len(missing), "ranges": ranges},
                     "所有有文字、域或对象的段落均有已确认角色或明确的保留审核记录",
                     "按 coverage.unclassified_paragraph_ids 逐段阅读并确认角色，尤其核对文末、附录和未识别的正文开端；空白段落除外。不要猜测为正文或标题，也不要把未检出格式问题当作已审核。",
                     "requires_role_review", "high")

    def check_paragraphs(self):
        editable_roles = defaultdict(list)
        for pid, (raw_role, confidence, basis) in self.roles.items():
            role = raw_role if raw_role in self.spec["roles"] else ROLE_ALIASES.get(raw_role, raw_role)
            if role not in self.spec["roles"]: continue
            p, pp = self.ps[pid], self.effective_p(self.ps[pid])
            target = self.spec["defaults"] | self.spec["roles"][role]
            repair = None
            if self.sid(p).startswith("FITRepair_"):
                from format_engine import _read_repair_style
                repair = _read_repair_style(self.styles.get(self.sid(p)))[1]
                for key, spec_key in (("before_pt","before_pt"),("after_pt","after_pt"),("page_break_before","page_break"),("keep_with_next","keep_next"),("keep_lines","keep_lines")):
                    if key in repair: target[spec_key] = repair[key]
                if "line_min_pt" in repair: target.update(line_pt=repair["line_min_pt"],line_rule="atLeast")
            native_front_layout = role.startswith("cover_") or role.startswith("commitment_")
            if raw_role in {"body_math", "table_math"} or p.findall(".//m:oMath", NS):
                target = target | {"line_rule": "atLeast"}
            mismatch = {}
            if role == "equation":
                if pvalue(pp, "jc", default="left") != "center" and not self.equation_tab_layout(p):
                    mismatch["alignment"] = {"actual": pvalue(pp, "jc", default="left"), "expected": "主体居中；普通段落也可用页宽中心/右边界制表位定位主体及标签"}
                line = number(pvalue(pp, "spacing", "line"))
                rule = pvalue(pp, "spacing", "lineRule", "auto")
                if rule != "atLeast" or line is None or line < 400:
                    mismatch["line_spacing"] = {"actual_twips": line, "actual_rule": rule,
                                                 "expected_minimum_pt": 20, "expected_rule": "atLeast"}
            # Cover align/spacing are actual-template layout, not universal overrides.
            if not native_front_layout and role != "equation":
                align = pvalue(pp, "jc", default="left")
                if align != target["align"]: mismatch["alignment"] = {"actual": align, "expected": target["align"]}
                line = pvalue(pp, "spacing", "line"); rule = pvalue(pp, "spacing", "lineRule", "auto")
                line_ok = number(line) is not None and (number(line) >= target["line_pt"] * 20 if target["line_rule"] == "atLeast" else number(line) == target["line_pt"] * 20)
                if not line_ok or rule != target["line_rule"]:
                    mismatch["line_spacing"] = {"actual_twips": line, "actual_rule": rule,
                                                 "expected_pt": target["line_pt"], "expected_rule": target["line_rule"]}
                for prop, spec_prop in (("before", "before_pt"), ("after", "after_pt")):
                    actual = pvalue(pp, "spacing", prop, "0")
                    # The required blank line before keywords may legitimately be
                    # represented as one body line of paragraph spacing.
                    keyword_gap = 20 if role == "cn_keywords" else 16 if role == "en_keywords" else None
                    gap_ok = prop == "before" and keyword_gap is not None and number(actual) == keyword_gap*20
                    # Table/figure separation may be realized with one body line
                    # of caption spacing rather than an empty paragraph.
                    caption_gap = (role == "table_caption" and prop == "before" or role == "figure_caption" and prop == "after" or role in {"figure", "source"} and prop in {"before", "after"}) and number(actual) == 400
                    gap_ok = gap_ok or caption_gap
                    if number(actual) != target[spec_prop] * 20 and not gap_ok:
                        mismatch["space_" + prop] = {"actual_twips": actual, "expected_pt": target[spec_prop]}
                first = pvalue(pp, "ind", "firstLineChars")
                if target.get("hanging_chars"):
                    hang = pvalue(pp, "ind", "hangingChars")
                    if number(hang) != target["hanging_chars"] * 100:
                        mismatch["hanging_indent"] = {"actual_chars100": hang, "expected_chars": target["hanging_chars"]}
                elif number(first) != target["indent_chars"] * 100:
                    # Equivalent fixed twips indentation may look identical, but no
                    # character-unit equality can be asserted without font/layout.
                    fixed = pvalue(pp, "ind", "firstLine", "0")
                    if not (target["indent_chars"] == 0 and number(first or "0") == 0 and number(fixed) == 0):
                        mismatch["first_line_indent"] = {"actual_chars100": first, "actual_twips": fixed,
                                                           "expected_chars": target["indent_chars"]}
                left = pvalue(pp, "ind", "leftChars")
                if target.get("left_chars") and number(left) != target["left_chars"] * 100:
                    mismatch["left_indent"] = {"actual_chars100": left, "actual_twips": pvalue(pp, "ind", "left"),
                                                "expected_chars": target["left_chars"]}
                outline = pvalue(pp, "outlineLvl", default="9")
                if str(target["outline"]) != outline:
                    mismatch["outline_level"] = {"actual": outline, "expected": target["outline"]}
                if target.get("page_break") and not bool_prop(pp, "pageBreakBefore"):
                    # A section break immediately before the heading also starts it
                    # on a new page; record the alternative instead of false failing.
                    if not self.section_break_before(p):
                        mismatch["new_page"] = {"actual": "无段前分页或紧邻新页分节", "expected": "另起页"}
                if target.get("keep_next") and not bool_prop(pp, "keepNext"):
                    mismatch["keep_next"] = {"actual": False, "expected": True}
                if target.get("keep_lines") and not bool_prop(pp, "keepLines"):
                    mismatch["keep_lines"] = {"actual": False, "expected": True}
            fonts = []
            if role != "equation":
                for ri, r in enumerate(p.iter(q("r")), 1):
                    if any(a.tag.startswith("{"+M+"}") for a in r.iterancestors()): continue
                    if next((a for a in r.iterancestors() if a.tag == q("p")), None) is not p: continue
                    t = text(r)
                    if not t.strip(): continue
                    rp = self.effective_r(p, r); differences = {}
                    for cat, pattern in (("east_asia", r"[\u3400-\u9fff]"), ("western", r"[A-Za-z0-9]")):
                        if re.search(pattern, t):
                            actual = self.font(rp, cat)
                            if actual != target[cat]: differences[cat] = {"actual": actual, "expected": target[cat]}
                    size = pvalue(rp, "sz")
                    if number(size) != target["size_pt"] * 2:
                        differences["size_pt"] = {"actual": number(size)/2 if number(size) is not None else None,
                                                  "expected": target["size_pt"]}
                    if differences: fonts.append({"run_index": ri, "text_excerpt": short(t, 32), "differences": differences})
            if fonts: mismatch["runs"] = {"count": len(fonts), "examples": fonts[:12]}
            if mismatch:
                self.add("warning", "paragraph.format", self.locate(pid), mismatch,
                         {"role": role, "style": target["name"]},
                         "确认语义角色后用对应命名样式统一有效字体、字号、段落属性；保留语义加粗、上下标和对象。",
                         "auto_fixable" if confidence == "high" else "requires_role_review", confidence)
            sid = self.sid(p); expected_id = "FIT_" + role
            # Different ID is allowed when the user's equivalent named style is
            # genuinely editable and already has the correct effective properties.
            st = self.styles.get(sid)
            name = val(st.find(q("name"))) if st is not None else None
            visible = st is not None and enabled(st.find(q("qFormat"))) and not enabled(st.find(q("semiHidden"))) and not enabled(st.find(q("hidden")))
            if st is None or not visible or enabled(st.find(q("locked"))) or enabled(st.find(q("autoRedefine"))):
                editable_roles[(role, sid, name)].append(pid)
            direct_p = props(p.find(q("pPr")))
            blockers = sorted(set(direct_p) & PARA_CONTROLS)
            if "ind" in blockers:
                ind = direct_p["ind"]
                dormant = {name for name in ("firstLine", "hanging", "left", "right")
                           if name in ind and pvalue(pp, "ind", name+"Chars") is not None}
                if ind and set(ind) <= dormant:
                    blockers.remove("ind")
                    self.evidence["inactive_indentation_fallback"].append(pid)
            mark_controls = sorted(set(props(p.find("w:pPr/w:rPr", NS))) & RUN_CONTROLS)
            direct_runs = []
            for ri, r in enumerate(p.iter(q("r")), 1):
                if next((a for a in r.iterancestors() if a.tag==q('p')),None) is not p: continue
                if not text(r).strip() or any(a.tag.startswith("{"+M+"}") for a in r.iterancestors()): continue
                controls = sorted(set(props(r.find(q("rPr")))) & RUN_CONTROLS)
                if controls: direct_runs.append({"run_index": ri, "properties": controls})
            if (blockers or direct_runs or mark_controls) and (sid.startswith("FIT_") or name and name.startswith("论文-")) and not native_front_layout and role != "equation":
                has_math = bool(p.findall(".//m:oMath", NS))
                nxt = p.getnext()
                while nxt is not None and nxt.tag == q("p") and not text(nxt).strip() and not nxt.findall(".//m:oMath", NS) and not nxt.findall(".//w:drawing", NS): nxt = nxt.getnext()
                followed_object = nxt is not None and (nxt.tag == q("tbl") or bool(nxt.findall(".//m:oMath", NS)) or bool(nxt.findall(".//w:drawing", NS)))
                contextual_layout = not direct_runs and not mark_controls and (
                    set(blockers) <= {"spacing"} and role in {"cn_keywords", "en_keywords", "table_caption", "figure_caption", "figure", "source"}
                    or has_math and set(blockers) <= {"spacing", "keepNext"}
                    or set(blockers) <= {"keepNext"} and (followed_object or role=="table_text" and any(a.tag==q("tbl") for a in p.iterancestors()))
                    or set(blockers) <= {"pageBreakBefore"} and role in {"chapter", "cn_title", "en_title", "toc_heading", "reference_heading", "acknowledgement_heading", "appendix_heading", "conclusion_heading"})
                self.add("info" if contextual_layout else "warning", "style.contextual_layout_override" if contextual_layout else "style.direct_format_override", self.locate(pid),
                         {"style_id": sid, "paragraph_properties": blockers,
                          "paragraph_mark_run_properties": mark_controls,
                          "run_override_count": len(direct_runs), "run_examples": direct_runs[:12]},
                         "常规字体、字号、行距、缩进等继承可见样式，用户修改样式后能同步变化",
                         "清理妨碍继承的直接字体/字号/段落格式；不要清除语义加粗、斜体、引用和数学上下标。" +
                         ("本处属于数学行高、对象引入段跟随、部分另页或图表留空的局部布局例外；保留合法设置或另设上下文样式，不应为统一而清除。" if contextual_layout else ""),
                         "requires_role_review" if contextual_layout else "auto_fixable")
        for (role, sid, name), pids in editable_roles.items():
            self.add("warning", "style.not_editable_in_gallery",
                     self.locate(pids[0], related_paragraph_ids=pids),
                     {"style_id": sid, "style_name": name, "paragraph_count": len(pids)},
                     {"named_role": role, "qFormat": True, "hidden": False, "locked": False},
                     "建立或修正可见、未锁定的语义段落样式，并让相关段落继承该样式。")

    def section_break_before(self, p):
        # Check direct body siblings only; a table's cell break is not a document break.
        prior = p.getprevious()
        while prior is not None and prior.tag == q("p"):
            sec = prior.find("w:pPr/w:sectPr", NS)
            if sec is not None and val(sec.find(q("type")), default="nextPage") in {"nextPage", "oddPage", "evenPage"}: return True
            if text(prior).strip(): break
            if prior.find(".//w:br[@w:type='page']", NS) is not None: return True
            prior = prior.getprevious()
        return False

    def check_structure(self):
        first_abstract = next((int(pid[1:]) for pid in self.evidence["cn_abstract_heading"]), None)
        region = list(self.ps.items())[:first_abstract-1] if first_abstract else list(self.ps.items())[:70]
        # The pledge itself names the school, so its text must not be used to
        # certify the preceding cover. The native template's school name is also
        # inside its logo image; exact media identity is legitimate evidence.
        pledge_pos = next((i for i,(_,p) in enumerate(region) if "诚信承诺书" in clean(text(p)) and len(clean(text(p))) < 100), len(region))
        cover_region = region[:pledge_pos]
        cover_text = "\n".join(text(p) for _, p in cover_region)
        cover_heading = bool(re.search(r"本科(?:生)?毕业论文", clean(cover_text)))
        logo_match = self.native_cover_media_match([p for _,p in cover_region])
        cover_paragraph_texts = [clean(text(p)) for _, p in cover_region]
        labels = {"title": r"^题目[:：]", "name": r"^(?:学生)?姓名[:：]", "student_id": r"^学号[:：]",
                  "college": r"^学院[:：]", "major": r"^专业[:：]", "year": r"^年级[:：]",
                  "supervisor": r"^指导教师(?:[（(]签名[）)])?[:：]", "date": r"^(?:时间|日期)[:：]"}
        label_evidence = {key: any(re.search(pattern, s) for s in cover_paragraph_texts) for key, pattern in labels.items()}
        cover_tables = {a for _,p in cover_region for a in p.iterancestors() if a.tag == q("tbl")}
        copy_mark = "（正本）" in clean(cover_text)
        missing_labels = [key for key, present in label_evidence.items() if not present]
        cover_school = cover_heading and logo_match and copy_mark and not missing_labels and bool(cover_tables)
        if not cover_school:
            plausible = cover_heading and all(label_evidence[key] for key in ("student_id", "college", "major", "supervisor"))
            self.add("warning" if plausible else "error", "front.school_cover_missing_or_unconfirmed", self.locate(region[0][0] if region else None),
                     {"school_name_found": "福州理工学院" in cover_text,
                      "native_template_cover_media_match": logo_match, "undergraduate_cover_heading_found": cover_heading,
                      "copy_mark_found": copy_mark, "field_label_evidence": label_evidence,
                      "missing_field_labels": missing_labels, "cover_table_ids": sorted(self.tid[t] for t in cover_tables if t in self.tid),
                      "cover_evidence_excerpt": short(cover_text.replace("\n", " "), 160)},
                     "学校原生校徽、论文/正本标识、信息填写表和题目/姓名/学号/学院/专业/年级/指导教师/时间字段",
                     "不能只凭写有校名认定采用学校封面。按范本确认校徽与填写表布局；移入原稿真实信息，空字段不猜填。校徽字节不同仅表示尚未识别，须渲染人工确认；新增固定组件须纳入允许变更范围。",
                     "requires_render_review" if plausible else "requires_content_authorization", "medium" if plausible else "high")
        pledge_ids = [pid for pid, p in self.ps.items() if "诚信承诺书" in clean(text(p)) and len(clean(text(p))) < 100]
        if cover_school and pledge_ids:
            from format_engine import front_template_geometry_issues
            geometry_issues = front_template_geometry_issues(self.doc,self.styles_root)
            if geometry_issues:
                self.add("warning", "front.template_geometry", self.locate(region[0][0] if region else None),
                         {"template_geometry_differences": geometry_issues},
                         "封面及承诺书的学校模板网格、字段位置和段落几何保持一致",
                         "从净化模板恢复对应格式并核对实际前两页；校徽一致不能证明布局正确。",
                         "requires_render_review")
        if not pledge_ids:
            self.add("error", "front.pledge_missing", {"part": "word/document.xml", "region": "封面之后、摘要之前"},
                     "未发现诚信承诺书标题", "学院本科毕业论文（设计）诚信承诺书原生结构",
                     "补入学院承诺书固定文字与表格并用原作者真实信息；签名和日期不得生成。只改格式授权不能补写这些字符。",
                     "requires_content_authorization")
        for role, expected in (("cn_abstract_heading", "中文摘要"), ("en_abstract_heading", "英文摘要 Abstract"),
                               ("cn_keywords", "中文关键词"), ("en_keywords", "英文关键词 Key words:"), ("toc_heading", "自动目录")):
            if not self.evidence[role]:
                self.add("error", "front.required_part_missing", {"part": "word/document.xml", "role": role},
                         "未识别到该部分；默认识别可能遗漏非常规样式", expected,
                         "先确认原稿是否已有该部分；不存在时报告缺失，不能以排版为由补写文章内容。",
                         "requires_content_authorization", "medium")
        seq = [(int(ids[0][1:]), role) for role, ids in self.evidence.items() if ids and role in {
            "cn_abstract_heading", "en_abstract_heading", "toc_heading", "reference_heading", "acknowledgement_heading", "appendix_heading"}]
        expected_order = ["cn_abstract_heading", "en_abstract_heading", "toc_heading", "reference_heading", "appendix_heading", "acknowledgement_heading"]
        ordered_roles = [r for _, r in sorted(seq)]
        if ordered_roles != [r for r in expected_order if r in ordered_roles]:
            self.add("warning", "structure.part_order", {"part": "word/document.xml"}, ordered_roles, expected_order,
                     "按已确认的学校要求重排后置部分：参考文献、已有附录、致谢；整段保留内容与对象。无附录不生成，不能为改顺序删除附录文字。", "auto_fixable", "medium")
        back_roles = {"chapter", "conclusion_heading", "reference_heading", "appendix_heading", "acknowledgement_heading"}
        structural = [(pid, role, confidence) for pid,(role,confidence,_) in self.roles.items() if role in back_roles]
        if not self.evidence["acknowledgement_heading"]:
            self.add("error", "structure.acknowledgement_missing_or_unconfirmed", {"part": "word/document.xml", "region": "文末"},
                     "未识别到正文中的致谢标题（目录缓存不作证据）", "最后一个实质部分为致谢；已有附录放其之前",
                     "先确认原稿是否已有致谢。存在时只重排并逐字保留；不存在时报告缺失，不补写致谢或生成附录。", "requires_content_authorization", "medium")
        elif structural and structural[-1][1] != "acknowledgement_heading":
            pid, role, confidence = structural[-1]
            self.add("error", "structure.acknowledgement_not_last", self.locate(pid),
                     {"last_identified_structure_role": role, "acknowledgement_heading_ids": self.evidence["acknowledgement_heading"]},
                     "最后已识别实质部分为致谢；已有附录位于致谢之前",
                     "按已确认的学校要求移动整段致谢至最后、已有附录至其前；保留全部段落、数据、公式、代码和对象，不删除或补写文章内容。",
                     "auto_fixable" if confidence == "high" else "requires_role_review", confidence)
        toc_fields = [f for f in fields(self.doc) if re.match(r"\s*TOC\b", f["instruction"], re.I)]
        if not toc_fields:
            self.add("error", "toc.no_real_field", {"part": "word/document.xml", "region": "目录"},
                     "未找到 TOC 域；可能为手打文字或只有 PAGEREF", '真实 TOC 域，默认 \\o "1-2"，两级',
                     "重建真实目录前需确认其标题及收录范围；域更新会改变目录文字/页码，不得冒充纯属性修改。",
                     "requires_content_authorization")
        for f in toc_fields:
            match = re.search(r'\\o\s+"([0-9]+)-([0-9]+)"', f["instruction"], re.I)
            styles_map = re.search(r'\\t\s+"([^"]+)"', f["instruction"], re.I)
            # A style-based TOC can be legitimate; disclose when depth is not
            # statically provable instead of reporting a false definite failure.
            if match and match.groups() != ("1", "2"):
                self.add("error", "toc.depth", {"part": "word/document.xml", "region": "目录域"}, f["instruction"],
                         '收录至两级（\\o "1-2" 或等效显式样式映射）',
                         "调整目录字段收录层级并按授权更新域；第三层现存条目不能静默删除。",
                         "requires_content_authorization")
            elif not match and not styles_map:
                self.add("warning", "toc.depth_unconfirmed", {"part": "word/document.xml", "region": "目录域"}, f["instruction"],
                         "可核对的两级收录规则", "在 Word 检查 TOC 字段代码与实际收录层级。", "requires_render_review")
        for pid in self.evidence["toc3"]:
            self.add("warning", "toc.third_level_entry", self.locate(pid), "三级目录样式条目", "学校目录只到二级",
                     "经目录重建/更新授权后排除第三层目录项，正文第三层标题继续保留。", "requires_content_authorization")
        self.check_toc_cached_numbering()

    def check_toc_cached_numbering(self):
        bookmarks = {}
        for marker in self.doc.iter(q("bookmarkStart")):
            paragraph = next((a for a in marker.iterancestors() if a.tag == q("p")), None)
            if paragraph in self.pid:
                bookmarks[val(marker, "name")] = self.pid[paragraph]
        for pid in self.evidence["toc1"] + self.evidence["toc2"]:
            paragraph = self.ps[pid]
            targets = []
            for record in fields(paragraph):
                match = re.match(r'\s*PAGEREF\s+(?:"([^"]+)"|([^\s\\]+))', record["instruction"], re.I)
                if match: targets.append(match.group(1) or match.group(2))
            targets += paragraph.xpath(".//w:hyperlink/@w:anchor", namespaces=NS)
            target = next((bookmarks[t] for t in targets if t in bookmarks and self.roles.get(bookmarks[t], (None,))[0] in {"chapter", "section"}), None)
            if target is None: continue
            target_role = self.roles[target][0]
            target_p = self.ps[target]; np = self.resolved_num(target_p)
            if np is None: continue
            level = val(np.find(q("ilvl")), default="0"); lvl = self.number_level(val(np.find(q("numId"))), level)
            expected = NUMBER_TARGETS[target_role]
            if lvl is None or (val(lvl.find(q("numFmt"))), val(lvl.find(q("lvlText")))) != expected[1:]: continue
            visible = cached_visible_text(paragraph).strip()
            decimal = re.match(r"^\d+(?:\.\d+)*(?:[.．]\s*|\s+)", visible)
            if decimal:
                self.add("warning", "toc.cached_number_prefix_mismatch", self.locate(pid, target_paragraph_id=target),
                         {"cached_prefix": decimal.group(), "target_role": target_role, "target_numId": val(np.find(q("numId"))), "target_lvlText": expected[2]},
                         "目录显示与其所指正文的学校中文编号一致；不能只改 TOC 代码而保留旧缓存",
                         "在已授权副本中更新/重建目录，再重新打开并核对中文章号、两级条目和实际页码；保留正文标题原文。", "requires_content_authorization")

    def native_cover_media_match(self, paragraphs):
        def image_hashes(parts, ps):
            rels = parse(parts["word/_rels/document.xml.rels"]) if "word/_rels/document.xml.rels" in parts else None
            targets = {r.get("Id"): posixpath.normpath(posixpath.join("word", r.get("Target", ""))) for r in rels if r.get("TargetMode") != "External"} if rels is not None else {}
            hashes = set()
            for p in ps:
                for el in p.iter():
                    rid = el.get("{"+R+"}embed") or el.get("{"+R+"}id")
                    path = targets.get(rid)
                    if path in parts and path.startswith("word/media/"): hashes.add(sha(parts[path]))
            return hashes
        template = Path(__file__).resolve().parent.parent / "assets" / "fit-economics-management-template.docx"
        if not template.exists(): return False
        with zipfile.ZipFile(template) as z: parts = {n:z.read(n) for n in z.namelist()}
        root = parse(parts["word/document.xml"]); cover_ps = []
        for p in root.iter(q("p")):
            if "诚信承诺书" in clean(text(p)): break
            cover_ps.append(p)
        return bool(image_hashes(self.parts, paragraphs) & image_hashes(parts, cover_ps))

    def check_numbering(self):
        numbering_ids = defaultdict(set)
        checked_style_bindings = set()
        for pid, (role, confidence, _) in self.roles.items():
            if role not in NUMBER_TARGETS: continue
            p = self.ps[pid]; np = self.resolved_num(p)
            expected_level, expected_format, expected_text = NUMBER_TARGETS[role]
            if np is None or val(np.find(q("numId"))) in {None, "0"}:
                typed = re.match(r"^\s*(?:第?[一二三四五六七八九十百]+[章、.]|[（(][一二三四五六七八九十0-9]+[）)]|\d+(?:[.、]\d+)*[.、]?)", text(p))
                self.add("error", "heading.not_true_automatic_numbering", self.locate(pid),
                         {"effective_numPr": None if np is None else E.tostring(np, encoding="unicode"),
                          "typed_number_prefix": typed.group() if typed else None},
                         {"ilvl": expected_level, "numFmt": expected_format, "lvlText": expected_text},
                         "绑定真实、同一四级列表。若原号是字符前缀，移除前缀并改自动编号会改变字符，须记录为授权的编号呈现迁移；不得重复显示两份编号。",
                         "requires_content_authorization" if typed else "auto_fixable", confidence)
                continue
            nid = val(np.find(q("numId"))); level = val(np.find(q("ilvl")), default="0")
            lvl = self.number_level(nid, level)
            actual = {"numId": nid, "ilvl": level, "numFmt": val(lvl.find(q("numFmt"))) if lvl is not None else None,
                      "lvlText": val(lvl.find(q("lvlText"))) if lvl is not None else None,
                      "lvlRestart": val(lvl.find(q("lvlRestart"))) if lvl is not None else None}
            numbering_ids[role].add(nid)
            if str(expected_level) != level or actual["numFmt"] != expected_format or actual["lvlText"] != expected_text:
                self.add("error", "heading.automatic_number_format", self.locate(pid), actual,
                         {"ilvl": expected_level, "numFmt": expected_format, "lvlText": expected_text},
                         "将四级标题绑定学院同一多级编号，子级随上级重启；保留标题文字，实际编号显示须在 Word 对照核验。",
                         "requires_content_authorization", confidence)
            elif (self.sid(p), nid, level) not in checked_style_bindings:
                checked_style_bindings.add((self.sid(p), nid, level))
                style_chain, seen, current = [], set(), self.sid(p)
                style_num = None
                while current in self.styles and current not in seen:
                    seen.add(current); style_chain.append(current)
                    candidate = self.styles[current].find("w:pPr/w:numPr", NS)
                    if style_num is None and candidate is not None: style_num = candidate
                    current = val(self.styles[current].find(q("basedOn")))
                inherited_binding = style_num is not None and val(style_num.find(q("numId"))) == nid and val(style_num.find(q("ilvl")), default="0") == level
                list_binding = val(lvl.find(q("pStyle"))) in style_chain
                if not inherited_binding and not list_binding:
                    self.add("warning", "heading.style_numbering_binding_unconfirmed", self.locate(pid),
                             {"style_id": self.sid(p), "numId": nid, "ilvl": level, "list_level_pStyle": val(lvl.find(q("pStyle"))), "style_chain": style_chain},
                             "标题样式与同一多级列表建立有效绑定，新增段落套样式仍能自动编号",
                             "现有段落有直接 numPr 不等于样式绑定完整。确认样式/列表关联，在内部副本中新增标题并检查真实编号；不要删改标题文字。", "requires_role_review", confidence)
        continuity_problems = self.heading_list_problems()
        if continuity_problems:
            self.add("warning", "heading.disconnected_lists", {"part": "word/numbering.xml"},
                     {"role_numIds": {k: sorted(v) for k, v in numbering_ids.items()},
                      "violations": continuity_problems},
                     "普通章节共用连续多级列表；已确认的无编号一级标题下使用独立、从1重启的受管子级列表",
                     "核对真实文档顺序及标题角色。正文不能章内换表；结论、附录等独立区块须各自绑定完整且唯一的重启定义，并渲染核对。", "requires_role_review")

    def heading_list_problems(self):
        """Accept only structurally justified, fully specified v10 restarts.

        A second numId alone is neither a defect nor proof of a valid restart.
        The boundary must be a confirmed unnumbered heading in the actual body,
        and all subordinate headings in that block must share its own list.
        """
        restart_roles = {"conclusion_heading", "reference_heading", "appendix_heading", "acknowledgement_heading"}
        body = self.doc.find(q("body")); blocks = []; active = None
        ordinary_ids, seen, problems = set(), set(), []

        def heading_id(pid):
            np = self.resolved_num(self.ps[pid])
            nid = val(np.find(q("numId"))) if np is not None else None
            return nid if nid not in {None, "0"} else None

        for p in body:
            pid = self.pid.get(p)
            role, confidence, _ = self.roles.get(pid, (None, None, None))
            if role in restart_roles:
                active = {"boundary": pid, "confirmed": confidence == "high",
                          "boundary_numId": heading_id(pid), "headings": []}
                blocks.append(active)
            elif role == "chapter":
                active = None
            if role not in NUMBER_TARGETS: continue
            seen.add(pid); nid = heading_id(pid)
            if active is None:
                if nid is not None: ordinary_ids.add(nid)
            else:
                active["headings"].append((pid, nid))
        # A heading in a table/container is not licensed by a nearby body
        # boundary. Retain the original conservative continuity requirement.
        for pid, (role, _, _) in self.roles.items():
            if role in NUMBER_TARGETS and pid not in seen:
                nid = heading_id(pid)
                if nid is not None: ordinary_ids.add(nid)
        if len(ordinary_ids) > 1:
            problems.append({"reason": "ordinary_headings_use_multiple_lists", "numIds": sorted(ordinary_ids)})

        def num_definition(nid):
            matches = self.numbering.findall(f"w:num[@w:numId='{nid}']", NS) if self.numbering is not None else []
            return matches[0] if len(matches) == 1 else None

        ordinary_aids = {val(n.find(q("abstractNumId"))) for nid in ordinary_ids
                         if (n := num_definition(nid)) is not None}
        claimed = set()
        for block in blocks:
            if not block["headings"]: continue
            ids = {nid for _, nid in block["headings"]}
            reasons = []
            if not block["confirmed"]: reasons.append("unconfirmed_boundary_role")
            if block["boundary_numId"] is not None: reasons.append("boundary_is_automatically_numbered")
            if len(ids) != 1 or None in ids:
                reasons.append("block_does_not_share_one_complete_list")
            else:
                nid = next(iter(ids)); num = num_definition(nid)
                if nid in ordinary_ids: reasons.append("block_reuses_ordinary_list")
                if nid in claimed: reasons.append("restart_list_reused_by_another_block")
                claimed.add(nid)
                aid = val(num.find(q("abstractNumId"))) if num is not None else None
                abstracts = self.numbering.findall(f"w:abstractNum[@w:abstractNumId='{aid}']", NS) if self.numbering is not None else []
                abstract = abstracts[0] if len(abstracts) == 1 else None
                if (abstract is None or val(abstract.find(q("name"))) != "FITv10Headings"
                        or ordinary_aids and ordinary_aids != {aid}):
                    reasons.append("restart_abstract_is_not_the_shared_managed_definition")
                else:
                    levels = abstract.findall(q("lvl"))
                    if len(levels) != 4 or {val(x, "ilvl") for x in levels} != {"0", "1", "2", "3"}:
                        reasons.append("managed_definition_has_incomplete_levels")
                    else:
                        for level, fmt, label in NUMBER_TARGETS.values():
                            lvl = abstract.find(f"w:lvl[@w:ilvl='{level}']", NS)
                            if (val(lvl.find(q("numFmt"))) != fmt or val(lvl.find(q("lvlText"))) != label
                                    or val(lvl.find(q("start"))) != "1"
                                    or val(lvl.find(q("lvlRestart"))) not in {None, str(level)}):
                                reasons.append("managed_definition_has_invalid_level_controls"); break
                overrides = num.findall(q("lvlOverride")) if num is not None else []
                if (len(overrides) != 3 or {val(x, "ilvl") for x in overrides} != {"1", "2", "3"}
                        or any(len(x) != 1 or x[0].tag != q("startOverride") or val(x[0]) != "1" for x in overrides)):
                    reasons.append("restart_requires_exact_child_level_start_overrides")
            if reasons:
                problems.append({"boundary": block["boundary"], "numIds": sorted(x for x in ids if x is not None),
                                 "headings": [pid for pid, _ in block["headings"]], "reasons": reasons})
        return problems

    def resolved_num(self, p):
        # Reuse the existing helper's robust inherited numPr/pStyle resolution.
        import thesis_format as F
        return F.resolved_numpr(p, self.styles, self.numbering)

    def number_level(self, nid, level):
        if self.numbering is None: return None
        num = self.numbering.find(f"w:num[@w:numId='{nid}']", NS)
        if num is None: return None
        over = num.find(f"w:lvlOverride[@w:ilvl='{level}']/w:lvl", NS)
        if over is not None: return over
        aid = val(num.find(q("abstractNumId")))
        return self.numbering.find(f"w:abstractNum[@w:abstractNumId='{aid}']/w:lvl[@w:ilvl='{level}']", NS)

    def check_sections(self):
        prior_end = 0
        inherited = {"header": {}, "footer": {}}
        rels = parse(self.parts["word/_rels/document.xml.rels"]) if "word/_rels/document.xml.rels" in self.parts else None
        relmap = {r.get("Id"): posixpath.normpath(posixpath.join("word", r.get("Target", "")))
                  for r in rels} if rels is not None else {}
        odd_even = enabled(self.settings.find(q("evenAndOddHeaders")))
        for i, sec in enumerate(self.doc.iter(q("sectPr")), 1):
            sid = f"s{i:03d}"
            parent = sec.getparent()
            end_p = parent.getparent() if parent is not None and parent.tag == q("pPr") else None
            end = int(self.pid[end_p][1:]) if end_p in self.pid else len(self.ps)
            start = prior_end + 1; prior_end = end
            roles = [self.roles[pid][0] for pid in list(self.ps)[start-1:end] if pid in self.roles]
            body = any(r in NUMBER_TARGETS or r in {"body", "body_math", "reference_heading", "reference", "appendix_heading", "acknowledgement_heading"} for r in roles)
            front = not body and any(r in {"cn_abstract", "cn_abstract_heading", "en_abstract", "en_abstract_heading", "toc1", "toc2", "toc3", "toc_heading"} for r in roles)
            kind = "body" if body else "front" if front else "unpaged"
            loc = self.locate(sid=sid, paragraph_range=[f"p{start:06d}", f"p{end:06d}"], section_role=kind)
            page_expected = self.spec["page"]
            page_actual = props(sec)
            differences = {}
            for attr, target in (("w", page_expected["width_twips"]), ("h", page_expected["height_twips"])):
                actual = pvalue(page_actual, "pgSz", attr)
                if number(actual) is None or abs(number(actual)-target) > 2:
                    differences["page_"+attr] = {"actual_twips": actual, "expected_twips": target}
            for attr in ("top", "bottom", "left", "right", "header", "footer"):
                actual = pvalue(page_actual, "pgMar", attr); target = page_expected[attr+"_twips"]
                if number(actual) is None or abs(number(actual)-target) > 2:
                    differences["margin_"+attr] = {"actual_twips": actual, "expected_twips": target}
            if pvalue(page_actual, "pgMar", "gutter", "0") != "0": differences["gutter"] = page_actual.get("pgMar")
            if differences: self.add("warning", "page.geometry", loc, differences, "A4 纵向；上/左 2.5 cm，下/右 2 cm；页眉页脚 1.5 cm；装订线 0",
                                     "按学校页面设置修正；原横向宽表节须先人工判断能否转为纵向。")
            refs, explicit = {}, {}
            for typ in ("header", "footer"):
                explicit[typ] = {}
                for ref in sec.findall(q(typ+"Reference")):
                    t = val(ref, "type", "default"); path = relmap.get(ref.get("{"+R+"}id"))
                    explicit[typ][t] = path; inherited[typ][t] = path
                refs[typ] = dict(inherited[typ])
            pg = sec.find(q("pgNumType"))
            num_format = val(pg, "fmt", "decimal"); num_start = val(pg, "start")
            info = {"id": sid, "kind": kind, "start_paragraph": f"p{start:06d}", "end_paragraph": f"p{end:06d}",
                    "page_number_format": num_format, "page_number_start": num_start,
                    "explicit_references": explicit, "effective_references": refs,
                    "title_page_different": enabled(sec.find(q("titlePg")))}
            self.sections.append(info)
            if kind in {"body", "front"}:
                expected_fmt = "decimal" if kind == "body" else "upperRoman"
                if num_format != expected_fmt:
                    self.add("error", "pagination.number_format", loc, {"fmt": num_format, "start": num_start},
                             {"fmt": expected_fmt}, "修复分节页码格式；前置摘要目录连续罗马号，正文从阿拉伯 1 连续至文末。")
                prior_same = [s for s in self.sections[:-1] if s["kind"] == kind]
                if not prior_same and num_start != "1":
                    self.add("error", "pagination.first_page_start", loc, {"fmt": num_format, "start": num_start},
                             {"start": "1", "display": "1" if body else "I"}, "在该部分首节设置从 1 开始；随后同部分各节续前节。")
                elif prior_same and num_start is not None:
                    self.add("error", "pagination.unwanted_restart", loc, {"start": num_start}, "该部分后续节连续，省略 start",
                             "去掉后续节重新编号，保持正文、参考文献、致谢和附录连续。")
                footer_types = ["default", "even"] if odd_even else ["default"]
                if info["title_page_different"]: footer_types.append("first")
                for t in footer_types:
                    path = refs["footer"].get(t)
                    ff = fields(parse(self.parts[path])) if path in self.parts else []
                    page_fields = [f for f in ff if re.match(r"\s*PAGE\b", f["instruction"], re.I)]
                    if not page_fields:
                        self.add("error", "pagination.no_page_field", self.locate(sid=sid, footer_type=t, part=path or "word/document.xml"),
                                 {"effective_footer_part": path, "field_instructions": [f["instruction"] for f in ff]},
                                 "该页脚有真实 PAGE 域，居中，TNR 9 磅", "修复所有实际生效的页脚；不得手打页码或只补奇数页。")
                    if path in self.parts: self.check_story_format(path, "page_number", sid, t)
                if body:
                    if not odd_even:
                        self.add("error", "header.odd_even_disabled", loc, False, "真实奇偶页不同开关及 default/even 页眉引用",
                                 "启用 Word 奇偶页不同，并在各正文节分别配置奇数题目和偶数固定学校文字。条件 IF 域不能替代本审核要求的真实奇偶页眉。")
                    for t, expected in (("default", self.title), ("even", SCHOOL_HEADER)):
                        path = refs["header"].get(t)
                        header_root = parse(self.parts[path]) if path in self.parts else None
                        actual = cached_visible_text(header_root) if header_root is not None else None
                        cond = [f["instruction"] for f in fields(header_root)] if header_root is not None else []
                        if actual is None or expected is not None and clean(actual) != clean(expected):
                            self.add("error", "header.wrong_effective_text", self.locate(sid=sid, header_type=t, part=path or "word/document.xml"),
                                     {"text": actual, "effective_part": path, "explicit_part": explicit["header"].get(t), "fields": cond},
                                     expected or "奇数页为已确认本论文完整题目", "使用本论文题目原文和指定学校名称，检查链接到前节及首页不同；按正文显示页码核验奇偶。")
                        if path in self.parts: self.check_story_format(path, "header", sid, t)
                    if info["title_page_different"]:
                        self.add("warning", "header.body_first_page_different", loc, refs["header"].get("first"),
                                 "正文第一页也使用与显示页码对应的页眉",
                                 "确认 first 页眉实际存在且正确；无特别需要时关闭正文首页不同。", "requires_render_review")
            else:
                for t, path in refs["footer"].items():
                    if path not in self.parts: continue
                    if any(re.match(r"\s*PAGE\b", f["instruction"], re.I) for f in fields(parse(self.parts[path]))):
                        self.add("warning", "pagination.cover_or_pledge_has_page_field", self.locate(sid=sid, footer_type=t, part=path),
                                 "有效页脚有 PAGE 域", "封面及承诺书无显示页码",
                                 "确认首页/奇偶页实际启用状态后隔离该节页脚，保留正文页码。", "requires_render_review")
        if not self.title:
            self.add("warning", "header.title_source_unconfirmed", {"part": "word/document.xml"}, None,
                     "已确认的本论文题目原文", "传 --title 或明确 cover_title/cn_title 角色定位，不能猜测或缩写题目。", "requires_role_review")

    def check_story_format(self, path, role, sid, typ):
        key = (path, role)
        if key in self.checked_story_formats: return
        self.checked_story_formats.add(key)
        root = parse(self.parts[path]); target = self.spec["defaults"] | self.spec["roles"][role]
        for index, p in enumerate(root.iter(q("p")), 1):
            if not cached_visible_text(p).strip(): continue
            pp = self.effective_p(p); mismatch = {}
            if pvalue(pp, "jc", default="left") != "center":
                mismatch["alignment"] = pvalue(pp, "jc", default="left")
            fonts = []
            for ri, r in enumerate(p.iter(q("r")), 1):
                t = text(r)
                if not t.strip(): continue
                rp = self.effective_r(p, r); diff = {}
                for cat, pattern in (("east_asia", r"[\u3400-\u9fff]"), ("western", r"[A-Za-z0-9]")):
                    if re.search(pattern, t) and self.font(rp, cat) != target[cat]:
                        diff[cat] = {"actual": self.font(rp, cat), "expected": target[cat]}
                size = number(pvalue(rp, "sz"))
                if size != target["size_pt"]*2:
                    diff["size_pt"] = {"actual": size/2 if size is not None else None, "expected": target["size_pt"]}
                if diff: fonts.append({"run_index": ri, "differences": diff})
            if fonts: mismatch["runs"] = fonts
            if role == "header":
                # Preserve the actual nested border node; flattened pPr attrs do
                # not carry pBdr/bottom's line width.
                border = None; chain = []; current = self.sid(p); seen = set()
                while current in self.styles and current not in seen:
                    seen.add(current); st = self.styles[current]; chain.append(st)
                    current = val(st.find(q("basedOn")))
                for st in reversed(chain):
                    b = st.find("w:pPr/w:pBdr/w:bottom", NS)
                    if b is not None: border = b
                b = p.find("w:pPr/w:pBdr/w:bottom", NS)
                if b is not None: border = b
                if border is None or val(border) != "single" or val(border, "sz") != "4":
                    mismatch["bottom_border"] = {"val": val(border), "eighth_pt": val(border, "sz")}
            if mismatch:
                self.add("warning", "header_footer.effective_format",
                         {"part": path, "section_id": sid, "story_paragraph_id": f"p{index:06d}", "story_type": typ},
                         mismatch, {"role": role, "font_east_asia": target["east_asia"], "font_western": target["western"],
                                    "size_pt": target["size_pt"], "alignment": "center", "header_bottom_border_pt": 0.5 if role == "header" else None},
                         "通过可见页眉/页码样式修复有效属性，并保留 PAGE 等字段结构。")

    def continuation_caption(self, table, label):
        """Recognize a managed split only through a compatible, captioned chain.

        The displayed words alone are not provenance. The format-stage verifier
        separately replays the source-bound split; this static audit also checks
        the actual adjacent tables and their repeated header structure.
        """
        body = self.doc.find(q("body"))
        mapped = self.mapping.get("tables", {})
        header_forbidden = {q(n) for n in (
            "tbl", "drawing", "pict", "object", "fldChar", "fldSimple", "instrText",
            "bookmarkStart", "bookmarkEnd", "ins", "del", "moveFrom", "moveTo",
            "commentRangeStart", "commentRangeEnd", "commentReference",
            "footnoteReference", "endnoteReference")}

        def signature(target):
            setting = mapped.get(self.tid.get(target), {})
            kind, nh = setting.get("kind"), setting.get("header_rows", 1)
            rows = target.findall(q("tr"))
            grid = target.find(q("tblGrid"))
            widths = tuple(val(col, "w") for col in grid) if grid is not None else ()
            if (kind not in {"three-line", "complex", "panel"} or type(nh) is not int
                    or not 0 < nh < len(rows) or not widths
                    or any(not (width or '').isdigit() or int(width) <= 0 for width in widths)):
                return None
            headers = []
            for ri, row in enumerate(rows):
                cells = row.findall(q("tc"))
                spans = [val(cell.find("w:tcPr/w:gridSpan", NS), default="1") for cell in cells]
                if (not cells or any(not span.isdigit() or int(span) <= 0 for span in spans)
                        or sum(map(int, spans)) != len(widths)):
                    return None
                if ri == nh and any(cell.find("w:tcPr/w:vMerge", NS) is not None
                                    and val(cell.find("w:tcPr/w:vMerge", NS), default="continue") != "restart"
                                    for cell in cells):
                    return None
                if ri >= nh:
                    continue
                if (not enabled(row.find("w:trPr/w:tblHeader", NS))
                        or any(el.tag in header_forbidden or el.tag.startswith("{" + M + "}") for el in row.iter())):
                    return None
                headers.append(tuple((span,
                    val(cell.find("w:tcPr/w:vMerge", NS), default="continue")
                        if cell.find("w:tcPr/w:vMerge", NS) is not None else None,
                    tuple(text(p) for p in cell.findall(q("p")))) for cell, span in zip(cells, spans)))
            if not any(text(row).strip() for row in rows[:nh]):
                return None
            return kind, nh, widths, tuple(headers)

        current, marker, links = table, label, []
        while marker is not None:
            if (current.getparent() is not body or marker.getparent() is not body
                    or marker.tag != q("p") or marker.getnext() is not current
                    or text(marker).strip() != "续前表" or self.sid(marker) != "FIT_continuation_label"):
                return None
            style = self.styles.get("FIT_continuation_label")
            pp = self.effective_p(marker)
            if (style is None or val(style, "type") != "paragraph"
                    or val(style.find(q("name"))) != self.spec["roles"]["continuation_label"]["name"]
                    or pvalue(pp, "jc") != "right" or pvalue(pp, "keepNext", default="0") in {"0", "false", "off"}
                    or marker.find("w:pPr/w:sectPr", NS) is not None
                    or any(child.tag != q("pPr") and child.tag != q("r") for child in marker)
                    or any(child.tag not in {q("rPr"), q("t")} for run in marker.findall(q("r")) for child in run)):
                return None
            previous = marker.getprevious()
            if previous is None or previous.tag != q("tbl"):
                return None
            shape = signature(current)
            if shape is None or shape != signature(previous):
                return None
            links.append({"table_id": self.tid[current], "preceding_table_id": self.tid[previous],
                          "label_paragraph_id": self.pid.get(marker)})
            before = previous.getprevious()
            while before is not None and before.tag == q("p") and not text(before).strip():
                before = before.getprevious()
            caption = text(before).strip() if before is not None and before.tag == q("p") else ""
            numbered = re.match(r"^(?:表\s*\d+\s*[-－—]\s*\d+|附表\s*\d+)", caption)
            if numbered and caption[numbered.end():].strip():
                return {"table_id": self.tid[table], "origin_table_id": self.tid[previous],
                        "origin_caption": caption, "links": links}
            current, marker = previous, before
        return None

    def check_tables(self):
        mapped = self.mapping.get("tables", {})
        for tid, table in self.ts.items():
            rows = table.findall(q("tr")); ps = [self.pid[p] for p in table.iter(q("p")) if p in self.pid]
            table_roles = {self.roles[x][0] for x in ps if x in self.roles}
            math_table = bool(table.findall(".//m:oMath", NS)) and max([len(r.findall(q("tc"))) for r in rows] or [0]) <= 2
            code_table = "appendix_code" in table_roles
            setting = mapped.get(tid)
            if setting is not None and setting.get("kind") in {"formula-layout", "equation", "cover", "pledge", "code"}: continue
            if setting is None and (math_table and len(rows) == 1 or code_table or table_roles & {"cover_info", "cover_title"}): continue
            if setting is None:
                if max([len(r.findall(q("tc"))) for r in rows] or [0]) < 2: continue
                self.add("warning", "table.type_unconfirmed", self.locate(tid=tid),
                         {"rows": len(rows), "max_cells": max([len(r.findall(q("tc"))) for r in rows] or [0])},
                         "数据表/复杂文字表/封面承诺表/公式布局表分别确认",
                         "在 role-map.tables 显式标注 kind 与 header_rows；未确认前不把表格统一套成三线表。", "requires_role_review", "low")
                continue
            kind, nh = setting.get("kind", "three-line"), setting.get("header_rows", 1)
            if not isinstance(nh, int) or nh < 0 or nh > len(rows): raise ValueError("表格映射 header_rows 非法："+tid)
            actual_borders = []
            for el in table.xpath("./w:tblPr/w:tblBorders/*|./w:tr/w:tc/w:tcPr/w:tcBorders/*|./w:tr/w:tblPrEx/w:tblBorders/*", namespaces=NS):
                if val(el) not in {"nil", "none", "0"}:
                    actual_borders.append({"side": local(el), "val": val(el), "sz_eighth_pt": val(el, "sz"),
                                           "xml_path": table.getroottree().getpath(el)})
            bad = []
            for b in actual_borders:
                if kind == "three-line" and b["side"] in {"left", "right", "insideH", "insideV"}: bad.append(b)
                if b["side"] in {"top", "bottom"} and b["sz_eighth_pt"] not in {"6", "12"}: bad.append(b)
            if kind == "three-line":
                # Verify physical cell-edge positions too; a six-point border on
                # every row is not a three-line table despite valid line widths.
                for ri, row in enumerate(rows):
                    for ci, cell in enumerate(row.findall(q("tc"))):
                        for edge in cell.findall("w:tcPr/w:tcBorders/*", NS):
                            if val(edge) in {"nil", "none", "0"}: continue
                            side = local(edge)
                            expected_sz = "12" if side == "top" and ri == 0 or side == "bottom" and ri == len(rows)-1 else "6" if side == "bottom" and ri == nh-1 or side == "top" and ri == nh else None
                            if side in {"top", "bottom"} and (expected_sz is None or val(edge, "sz") != expected_sz or val(edge) != "single"):
                                bad.append({"row_index": ri+1, "cell_index": ci+1, "side": side,
                                            "actual_sz": val(edge, "sz"), "expected_sz": expected_sz})
                table_borders = table.find("w:tblPr/w:tblBorders", NS)
                top_ok = table_borders is not None and val(table_borders.find(q("top"))) == "single" and val(table_borders.find(q("top")), "sz") == "12"
                bottom_ok = table_borders is not None and val(table_borders.find(q("bottom"))) == "single" and val(table_borders.find(q("bottom")), "sz") == "12"
                if not top_ok and not any(val(c.find("w:tcPr/w:tcBorders/w:top", NS)) == "single" and val(c.find("w:tcPr/w:tcBorders/w:top", NS), "sz") == "12" for c in (rows[0].findall(q("tc")) if rows else [])):
                    bad.append({"required_line": "top", "expected_eighth_pt": "12", "found": False})
                if not bottom_ok and not any(val(c.find("w:tcPr/w:tcBorders/w:bottom", NS)) == "single" and val(c.find("w:tcPr/w:tcBorders/w:bottom", NS), "sz") == "12" for c in (rows[-1].findall(q("tc")) if rows else [])):
                    bad.append({"required_line": "bottom", "expected_eighth_pt": "12", "found": False})
                if nh and nh <= len(rows):
                    middle_ok = any(val(c.find("w:tcPr/w:tcBorders/w:bottom", NS)) == "single" and val(c.find("w:tcPr/w:tcBorders/w:bottom", NS), "sz") == "6" for c in rows[nh-1].findall(q("tc")))
                    if not middle_ok and nh < len(rows):
                        middle_ok = any(val(c.find("w:tcPr/w:tcBorders/w:top", NS)) == "single" and val(c.find("w:tcPr/w:tcBorders/w:top", NS), "sz") == "6" for c in rows[nh].findall(q("tc")))
                    if not middle_ok: bad.append({"required_line": "header_bottom", "expected_eighth_pt": "6", "found": False})
            tp = props(table.find(q("tblPr")))
            if pvalue(tp, "jc", default="left") != "center" or bad or not actual_borders:
                self.add("warning", "table.borders_or_alignment", self.locate(tid=tid),
                         {"alignment": pvalue(tp, "jc", default="left"), "bad_borders": bad[:20], "visible_border_count": len(actual_borders)},
                         "居中；顶底 1.5 磅；三线表仅表头下 0.75 磅，复杂文字表内线 0.75 磅，无左右外框",
                         "依已确认表格类型清理表/单元格/条件样式中的冲突边线；复核表头分隔线位置。")
            for ri, row in enumerate(rows):
                tr = row.find(q("trPr"))
                if ri < nh and not enabled(tr.find(q("tblHeader")) if tr is not None else None):
                    self.add("warning", "table.repeat_header_missing", self.locate(tid=tid, row_index=ri+1), False,
                             "已确认表头行启用跨页重复标题行", "打开重复表头；跨页是否显示及续表标识须渲染复核。")
                height = tr.find(q("trHeight")) if tr is not None else None
                if height is not None and val(height, "hRule") == "exact":
                    self.add("warning", "table.exact_row_height", self.locate(tid=tid, row_index=ri+1),
                             dict(height.attrib), "不固定会截字的行高", "去掉固定行高或改最小高度，并检查长内容跨页。")
            prior = table.getprevious()
            while prior is not None and prior.tag == q("p") and not text(prior).strip(): prior = prior.getprevious()
            caption = text(prior).strip() if prior is not None and prior.tag == q("p") else ""
            if not re.match(r"^(?:表\s*\d+\s*[-－—]\s*\d+|附表\s*\d+\s)", caption):
                continuation = self.continuation_caption(table, prior) if caption == "续前表" else None
                if continuation is not None:
                    self.evidence["table_continuations"].append(continuation)
                else:
                    self.add("warning", "table.caption_missing_or_position", self.locate(tid=tid),
                             {"preceding_text": short(caption)}, "表序与表名在表上方居中，按章编号；续表须有受管标记、相容的前表和重复表头",
                             "确认表题位置；续前表须追溯到原表题并核对列结构和重复表头，不能仅凭这三个字免除审核。", "requires_content_authorization", "medium")
        if self.ts:
            self.add("info", "table.render_gate", {"part": "word/document.xml", "table_ids": list(mapped)},
                     "XML 不包含最终跨页布局", "表不越页边、无截字，跨页有重复表头；需要时续前表标識",
                     "逐页查看 Word/PDF；代码和公式布局表不得按数据表规则强拆。", "requires_render_review")

    def check_math(self):
        math_count, ole_count, image_candidates = 0, 0, []
        formula_units = []
        for pid, p in self.ps.items():
            omml = p.findall(".//m:oMath", NS); objects = p.findall(".//w:object", NS)
            equation_ole = []
            for obj in objects:
                for ole in obj.findall(".//o:OLEObject", NS):
                    prog = ole.get("ProgID", "")
                    if re.search(r"equation|mathtype|dsmt", prog, re.I): equation_ole.append(prog)
            math_count += len(omml); ole_count += len(equation_ole)
            role = self.roles.get(pid, (None, None, None))[0]
            display = bool(p.findall(".//m:oMathPara", NS)) or role in {"equation", "equation_number"}
            table = next((a for a in p.iterancestors() if a.tag == q("tbl")), None)
            if table is not None and (omml or equation_ole) and max([len(r.findall(q("tc"))) for r in table.findall(q("tr"))] or [0]) <= 2:
                display = True
            if table is not None:
                tid = self.tid[table]
                table_roles = {self.roles.get(self.pid.get(tp), (None, None, None))[0] for tp in table.iter(q("p"))}
                mapped_kind = self.mapping.get("tables", {}).get(tid, {}).get("kind")
                is_data_table = mapped_kind in {"three-line", "complex", "data"} or bool(table_roles & {"table_text", "table_math"})
                if is_data_table:
                    display = False
                elif (len(table.findall(q("tr"))) != 1 or not table_roles & {"equation", "equation_number"}) and mapped_kind not in {"formula-layout", "equation"}:
                    # Two-column symbol/data tables can contain many OMML tokens;
                    # do not label them a definite display equation by width alone.
                    if omml or equation_ole:
                        self.evidence["formula_table_candidates"].append(tid)
                    display = False
            if display and (omml or equation_ole):
                unit = table if table is not None else p
                if not any(unit is old for old in formula_units): formula_units.append(unit)
            if role in {"equation", "equation_number"} and not omml and not equation_ole:
                if p.findall(".//w:drawing", NS) or p.findall(".//w:pict", NS): image_candidates.append(pid)
                elif re.search(r"[=＝+＋−×÷∑∫]|[_^]", text(p)):
                    self.add("error", "math.plain_text_formula", self.locate(pid), short(text(p)),
                             "公式编辑器 OMML，或经原生编辑确认可打开的旧式 Equation/MathType 对象",
                             "先确认数学含义，再按授权转成编辑器对象；不能凭字符串自动猜上下标或运算。", "requires_content_authorization")
            if role in {"body", "body_math", "cn_abstract", "en_abstract", "table_text", "table_math"} and not equation_ole:
                candidates = plain_math_candidates(p)
                if candidates:
                    self.add("warning", "math.plain_text_candidate", self.locate(pid),
                             {"ordinary_text_math_candidates": candidates, "existing_omml_count": len(omml)},
                             "已确认的公式与数学变量使用可编辑数学对象；普通下划线不冒充数学下标",
                             "先确认这些普通文字是否公式/变量或代码、路径、数据。确认后逐项按授权转为编辑器对象并对照原符号和上下标；本审核不自动转换、不猜数学含义。", "requires_role_review", "medium")
            if equation_ole:
                self.add("info", "math.legacy_equation_object", self.locate(pid), {"ProgID": equation_ole},
                         "可实际双击编辑的公式对象", "存在旧 OLE 对象只能证明嵌入类型；在 Word 打开确认编辑器可用。不要把预览图片当作已可编辑。", "requires_render_review")
            if (omml or equation_ole) and display:
                pp = self.effective_p(p)
                if pvalue(pp, "spacing", "lineRule", "auto") == "exact":
                    self.add("warning", "math.fixed_line_height", self.locate(pid),
                             {"line_twips": pvalue(pp, "spacing", "line"), "rule": "exact"},
                             "公式段至少 20 磅且容纳分式、根式及上下限；范本未限定内部数学字号",
                             "改合适的最小行距并渲染确认无裁切；内部数学字体/字号不得盲目统一。")
        for unit in formula_units:
            pid = self.pid.get(unit); tid = self.tid.get(unit)
            s = text(unit)
            label = re.search(r"公式\s*（\s*\d+\s*[-－—]\s*\d+\s*）", s)
            old = re.search(r"[（(]\s*\d+\s*[-－—]\s*\d+\s*[）)]", s)
            fs = fields(unit)
            seq = [f["instruction"] for f in fs if re.match(r"\s*SEQ\b", f["instruction"], re.I)]
            loc = self.locate(pid=pid, tid=tid)
            unit_pids = [self.pid[p] for p in unit.iter(q("p")) if p in self.pid] if tid else [pid]
            confirmed = any(self.roles.get(x, (None, None, None))[0] in {"equation", "equation_number"}
                            and self.roles[x][1] == "high" for x in unit_pids)
            formula_confidence = "high" if confirmed else "medium"
            layout_status = "auto_fixable" if confirmed else "requires_role_review"
            if tid:
                self.add("error", "math.table_layout_disallowed", loc,
                         {"table_id": tid, "rows": len(unit.findall(q("tr")))},
                         "独立公式用普通段落中的原生 OMML 与完整动态标签；不使用表格或文本框",
                         "将已审定公式主体、标签、SEQ/REF 和书签原样移到普通段落，用中心/右制表位排版；去除空布局容器属于格式迁移，不删除或改写数学内容。",
                         layout_status, formula_confidence)
            textbox = any(local(a) == "txbxContent" for a in unit.iterancestors()) or any(local(e) == "txbxContent" for e in unit.iter() if isinstance(e.tag,str))
            if textbox:
                self.add("error", "math.textbox_layout_disallowed", loc, "独立公式位于或包含文本框",
                         "普通正文段落中的原生公式对象，不使用文本框",
                         "定点移出已审定公式及完整字段/书签到普通段落；保留所有文字和数学对象，复核锚点及阅读顺序。",
                         layout_status, formula_confidence)
            visible_borders = []
            if tid:
                border_nodes = unit.xpath("./w:tblPr/w:tblBorders/*|./w:tr/w:tc/w:tcPr/w:tcBorders/*|./w:tr/w:tblPrEx/w:tblBorders/*", namespaces=NS)
            else:
                pb = self.effective_paragraph_container(unit, "pBdr")
                border_nodes = list(pb) if pb is not None else []
            for edge in border_nodes:
                if val(edge) not in {None, "nil", "none", "0"}:
                    visible_borders.append({"side": local(edge), "val": val(edge), "size_eighth_pt": val(edge, "sz")})
            if visible_borders:
                self.add("warning", "math.visible_border", loc, visible_borders,
                         "公式主体与标签无可见边框；普通段落布局",
                         "清除已确认公式布局的可见边框，保留公式及标签；同时移出被用户禁用的表格/文本框结构。Word 非打印虚线网格另按布局容器检查，不能声称它是打印边框。",
                         layout_status, formula_confidence)
            if not label:
                self.add("error", "math.label_incomplete_or_missing", loc,
                         {"number_label_found": old.group() if old else None, "unit_text_excerpt": short(s)},
                         "完整右侧题注：公式（x-x），主体居中；公式编号独立按章递增",
                         "只有（x-x）须补“公式”标签；缺编号需按授权建立。新增标签/改变现有编号属于允许变更范围，必须单列核验。",
                         "requires_content_authorization" if confirmed else "requires_role_review", formula_confidence)
            if not seq:
                self.add("warning", "math.static_number", loc, {"field_instructions": [f["instruction"] for f in fs]},
                         "动态 SEQ 编号与章序来源，正文需要引用时用带书签的 REF",
                         "将已确认公式编号迁移为动态字段并核验显示不重号；不得自动更新全篇域改变未授权文字。",
                         "requires_content_authorization" if confirmed else "requires_role_review", formula_confidence)
            elif not any(re.match(r"\s*STYLEREF\b", f["instruction"], re.I) for f in fs) and len(seq) < 2:
                self.add("warning", "math.chapter_number_not_dynamic", loc,
                         {"field_instructions": [f["instruction"] for f in fs]},
                         "公式编号章序来源也动态关联章标题/章节序列，序号按章重启",
                         "核对章序是否是固定文字；按已有编号迁移授权修复章序来源并记录字段显示结果。",
                         "requires_content_authorization" if confirmed else "requires_role_review", formula_confidence)
        for pid in image_candidates:
            self.add("error", "math.image_only_candidate", self.locate(pid),
                     "公式角色段落有图片而无 OMML/公式 OLE", "真实可编辑公式对象",
                     "人工辨认图片含义和上下标后才能重新录入；仅图片不能声明公式编辑器格式已达标。", "requires_content_authorization", "medium")
        malformed = [f for f in fields(self.doc) if re.match(r"\s*(?:SEQ|REF|STYLEREF)\b", f["instruction"], re.I) and not f["complete"]]
        if malformed:
            self.add("error", "fields.incomplete", {"part": "word/document.xml"}, [f["instruction"] for f in malformed],
                     "字段 begin/separate/end 完整，SEQ/REF 可解析", "定点修复字段结构，不全篇刷新域。")
        bookmarks = {val(b, "name") for b in self.doc.iter(q("bookmarkStart"))}
        for pid, p in self.ps.items():
            for f in fields(p):
                match = re.match(r'\s*REF\s+(?:"([^"]+)"|([^\s\\]+))', f["instruction"], re.I)
                if match and (match.group(1) or match.group(2)) not in bookmarks:
                    self.add("error", "fields.ref_bookmark_missing", self.locate(pid), f["instruction"],
                             "REF 指向本稿存在的正确书签", "确认交叉引用对象后修复书签目标，不能猜连另一公式或全篇刷新域。", "requires_role_review")
        self.add("info", "math.manual_semantics_gate", {"part": "word/document.xml"},
                 {"omml_count": math_count, "legacy_equation_ole_count": ole_count, "display_formula_units": len(formula_units),
                  "unconfirmed_formula_table_ids": sorted(set(self.evidence["formula_table_candidates"]))},
                 "所有公式和正文变量用编辑器，上下标、幂次、上下限及引用对象含义正确",
                 "OOXML 可识别对象和 sSub/sSup 等结构，不能可靠判断普通字母是否数学变量，也不能证明上下标的学术含义正确；逐项对原稿人工核查。", "requires_render_review")

    def check_citations(self):
        from narrative_citations import inspect_parts
        narrative = inspect_parts(self.parts)
        matches = defaultdict(list)
        for match in narrative.get("matches", narrative["candidates"]):
            matches[match["pid"]].append(match)
        self.citation_review = {
            "bibliography_entry_count": len(narrative["references"]),
            "matched_author_year_count": sum(len(values) for values in matches.values()),
            "candidate_count": len(narrative["candidates"]),
            "automatic_reference_assignment": False,
            "notes": narrative["notes"],
        }
        # Bibliography listing differs from the narrative display convention.
        # Inspection is advisory; never shorten a real reference entry here.
        for reference in narrative["references"]:
            authors = reference.get("authors", [])
            truncated = reference.get("truncation")
            problems = []
            if len(authors) > 3 and not truncated:
                problems.append("more_than_three_authors_listed_without_truncation")
            if truncated and len(authors) != 3:
                problems.append("truncation_does_not_follow_three_listed_authors")
            area = reference.get("author_list", {}).get("expected_text", "")
            if len(authors) > 1 and re.search(r"[，、;；&]|\band\b|和|与", area):
                problems.append("author_separator_is_not_an_English_comma")
            if problems:
                self.add("warning", "reference.author_listing_needs_review", self.locate(reference["pid"]),
                         {"author_list": area, "listed_author_count": len(authors),
                          "truncation": truncated, "problems": problems},
                         "文末三位以内全列，超过三位列前三位后加等或 et al.，作者之间用英文逗号",
                         "对照完整原条目查证作者名单；严格内容模式只报告，不能套正文三人简称规则或自动改文末条目。",
                         "requires_content_authorization", "low")
        expectations = {
            "missing_year": (
                "叙述作者后应有与已确认书目一致的出版年份",
                "核对本稿已有条目与引用范围后，用来源绑定记录补入年份；不猜测或改写书目年份。"),
            "missing_numeric_marker": (
                "已确认引用在所指内容末尾、句号前带与本稿书目一致的完整 [n] 上标",
                "核对完整论述与已有书目，确定引用范围及号码后记录来源映射；纯格式模式只报告，不自动补号。"),
            "author_count_needs_review": (
                "本包统一写法：一位写该作者；两位中英文均用和；三位及以上中文首作者加等、英文首作者加 et al.",
                "核对文末已列作者；正文惯例与文末三位以内全列规则不同。已有引用修正授权时按来源映射改，纯格式模式保留字符。"),
            "reference_ambiguous": (
                "每处引用对应经过确认的现有完整书目条目",
                "同作者同年份或多条书目不得自动选第一条；确认完整条目及引用内容后绑定准确书目 PID。"),
            "numeric_marker_needs_review": (
                "正文已写号码与所述作者年份对应的现有书目一致",
                "已有数字号不等于引用正确；核对号码、引用范围及完整书目，错号或多文献并引保留人工复核，不自动改号。"),
        }
        for candidate in narrative["candidates"]:
            pid = candidate["pid"]
            role = self.roles.get(pid, (None,))[0]
            if role not in {"body", "body_math", "cn_abstract", "en_abstract"}:
                continue
            for issue in candidate["issues"]:
                if issue not in expectations:
                    continue
                expected, recommendation = expectations[issue]
                self.add("warning", "citation." + issue,
                         self.locate(pid, character_span=[candidate["start"], candidate["end"]]),
                         {"author_year": candidate["expected_text"],
                          "reference_pids": candidate["reference_pids"],
                          "eligible_structure": candidate["eligible_structure"]},
                         expected, recommendation, "requires_role_review", "low")
        token = re.compile(r"\[\s*\d+(?:\s*[,，–—-]\s*\d+)*\s*\]")
        for pid, (role, confidence, _) in self.roles.items():
            p = self.ps[pid]
            if role not in {"body", "body_math", "cn_abstract", "en_abstract", "reference"}: continue
            runs = []
            for ri, r in enumerate(p.iter(q("r")), 1):
                if any(a.tag.startswith("{"+M+"}") for a in r.iterancestors()): continue
                t = text(r)
                if t: runs.append((ri, t, self.effective_r(p, r)))
            full = "".join(t for _, t, _ in runs)
            for match in token.finditer(full):
                lo, hi = match.span(); pos = 0; effects = []
                for ri, t, rp in runs:
                    if pos < hi and pos+len(t) > lo:
                        effects.append({"run_index": ri, "vertAlign": pvalue(rp, "vertAlign", default="baseline")})
                    pos += len(t)
                expected = "baseline" if role == "reference" else "superscript"
                if any(e["vertAlign"] != expected for e in effects):
                    self.add("warning", "citation.number_position", self.locate(pid, character_span=[lo, hi]),
                             {"token": match.group(), "runs": effects}, expected,
                             "正文确认是文献编号后对完整方括号施加引用上标；文末著录编号保持基线。数组/数据/代码不能按文献处理。",
                             "auto_fixable" if confidence == "high" and role == "reference" else "requires_role_review", confidence)
            # Character spans cover split Word runs, including the brackets.
            baseline_spans = [(m.start(), m.end(), "citation.year_not_baseline")
                              for m in re.finditer(r"[（(]\s*(?:19|20)\d{2}[a-z]?\s*[）)]", full)]
            baseline_spans += [(m["start"], m["end"], "citation.author_year_not_baseline")
                               for m in matches[pid]]
            for lo, hi, category in baseline_spans:
                pos, effects = 0, []
                for ri, t, rp in runs:
                    va = pvalue(rp, "vertAlign", default="baseline")
                    position = pvalue(rp, "position", default="0")
                    if pos < hi and pos + len(t) > lo and (va != "baseline" or position not in {"0", None}):
                        effects.append({"run_index": ri, "vertAlign": va, "position": position})
                    pos += len(t)
                if effects:
                    self.add("warning", category, self.locate(pid, character_span=[lo, hi]),
                             {"text": full[lo:hi], "runs": effects}, "作者姓名、连接词和括号内年份保持正文基线",
                             "确认作者年份引用后仅修复对应字符格式；完整数字引用方括号保留上标，公式、脚注和数学下标不得误清除。",
                             "requires_role_review")
            if role in {"body", "body_math"} and re.search(r"公式\s*（\s*\d+\s*[-－—]\s*\d+\s*）", full):
                ff = fields(p)
                if not any(re.match(r"\s*REF\b", f["instruction"], re.I) for f in ff):
                    self.add("warning", "math.body_reference_not_dynamic", self.locate(pid),
                             {"reference_text_excerpt": short(full), "field_instructions": [f["instruction"] for f in ff]},
                             "确认是公式交叉引用后使用真实书签 REF 字段",
                             "先确认该处引用对象；按已有授权把公式引用迁移为 REF，保留显示文本并核验书签目标。", "requires_role_review")

    def check_review_hygiene(self):
        """Detect inherited template notes without deleting manuscript reviews."""
        self.review_hygiene = report = inspect_review_parts(self.parts)
        if report["template_comment_matches"]:
            self.add("warning", "review.template_comment_match", {"part": "DOCX package"},
                     report["template_comment_matches"], "成稿不携带学院范本的历史教学批注",
                     "逐条核实来源；复用学校组件时使用净化模板。文字相同是来源线索，不能据此自动删除作者或导师的真实批注。",
                     "requires_role_review")
        orphans = {key: report[key] for key in ("orphan_anchor_ids", "missing_id_anchors",
                   "dangling_comment_relationships", "dangling_comment_content_types") if report[key]}
        if orphans:
            self.add("error", "review.dangling_comment_metadata", {"part": "DOCX package"},
                     orphans, "批注锚点、文件关系及内容类型声明均指向存在的部件和批注",
                     "在副本中核对并修复孤立批注登记，记录实际变化；不得接受或删除修订，不得删除承载图片或域的整个 run。",
                     "requires_role_review")
        unknown_count = report["comment_count"] - len(report["template_comment_matches"])
        if unknown_count:
            self.add("info", "review.manuscript_comments_preserved", {"part": "DOCX package"},
                     {"comment_count_without_template_text_match": unknown_count},
                     "作者或导师批注受保护；审核说明与论文分开交付",
                     "仅记录批注存在，不推定其作者或来源；没有明确的删除要求时保留。", "requires_role_review")
        revisions = {key: count for key, count in report["revisions"].items() if count}
        if revisions:
            self.add("info", "review.tracked_revisions_preserved", {"part": "DOCX package"},
                     revisions, "精确按 XML 名称清点修订，批注清理不改变修订",
                     "保留原稿真实修订；w:insideH 和 w:instrText 不属于 w:ins。计数是 XML 节点数，不是人工编辑次数。",
                     "requires_role_review")

    def run(self):
        self.infer_roles()
        self.check_coverage()
        self.check_review_hygiene()
        self.check_structure(); self.check_paragraphs(); self.check_numbering()
        self.check_sections(); self.check_tables(); self.check_math(); self.check_citations()
        protection = self.settings.find(q("documentProtection"))
        if protection is not None and enabled_attr(protection, "enforcement"):
            self.add("warning", "document.edit_protection", {"part": "word/settings.xml"},
                     {local(k): v for k, v in protection.attrib.items()}, "输出可编辑，已有保护限制如实说明",
                     "确认保护类型与权限；有保护不能声称用户可自由手调所有样式。", "requires_render_review")
        if self.source.read_bytes() != self.before: raise RuntimeError("审核期间源文件发生变化；本次报告无效。")
        severities = Counter(i["severity"] for i in self.issues)
        statuses = Counter(i["status"] for i in self.issues)
        return {
            "schema_version": 1, "mode": "read_only_format_audit", "created_at_utc": datetime.now(timezone.utc).isoformat(),
            "source": str(self.source), "source_sha256": self.hash, "source_unchanged": True,
            "style_spec_sha256": sha(self.spec_bytes), "title_used_for_header_check": self.title,
            "summary": {"paragraph_count": len(self.ps), "table_count": len(self.ts), "section_count": len(self.sections),
                        "classified_paragraph_count": len(self.roles), "issue_count": len(self.issues),
                        "substantive_paragraph_count": self.coverage["substantive_paragraph_count"],
                        "unclassified_substantive_paragraph_count": len(self.coverage["unclassified_paragraph_ids"]),
                        "by_severity": dict(severities), "by_status": dict(statuses),
                        "machine_format_result": "issues_found" if severities["error"] or severities["warning"] else "no_machine_detected_issue",
                        "render_review_required": True, "not_a_final_compliance_certificate": True},
            "role_identification": {pid: {"role": role, "confidence": c, "basis": b} for pid, (role, c, b) in self.roles.items()},
            "coverage": self.coverage,
            "citation_review": self.citation_review,
            "review_hygiene": self.review_hygiene,
            "sections": self.sections, "issues": self.issues,
            "style_layout_exceptions": {
                "native_front_layout_paragraph_ids": [pid for pid,(role,_,_) in self.roles.items() if role.startswith("cover_") or role.startswith("commitment_")],
                "inactive_twips_indentation_fallback_ids": self.evidence["inactive_indentation_fallback"],
                "note": "原生封面/承诺书布局可保留局部属性；继承字符单位缩进时 Word 保存的 twips 后备值不生效，不算样式阻断。公式最小行高及必要对象跟随另列 info。",
            },
            "limitations": [
                "定位是 DOCX XML 段落/表格/节 ID，不是打印页面编号。没有渲染就不能证明页眉实际出现在正确奇偶页、没有空白页或公式未裁切。",
                "未传显式 SHA-256 绑定映射时，角色按命名样式、大纲级别和有限文本模式识别；中低置信度项目须先人工确认。普通数字开头段落不据此认定为标题。",
                "未确认类型的表格只报告待分类，封面/承诺表/公式布局/代码表不能统一套三线表。已映射表格仍需渲染确认边线和跨页布局。",
                "正文变量与数学含义、图片内字体、旧式 OLE 是否能实际编辑、公式上下标是否符合原意不能仅凭 XML 保证。",
                "摘要页数、目录页码同步、标题悬孤、图表重叠和越界、跨页续表及签名日期须结合原生 Word/PDF 检查。",
                "未执行文字修改、目录/编号/公式字段刷新或页面重排；报告不会改变论文。公式/标题编号迁移、新封面承诺书及缺失部分补写需已有明确授权。",
                "引用模式可能对应数组/数据等，须核实为文献后再改上标；本工具不编造作者、年份、文献数量或缺失内容。",
                "作者年份提示只匹配本稿可识别的书目姓名与年份，可能漏掉不支持的外文姓名、未列书目或其他写法；候选提示不是引用确认，须逐处审核全文。",
                "致谢最后的静态判断依据已识别的实质结构标题，不能证明未识别散文、文本框或实际最后打印页归属；最终须看渲染的最后一页。",
                "公式表格/文本框是本用户禁止的布局容器，结构迁移不等于删除学术内容。静态边框检查覆盖明确的表/单元格/段落边框，主题条件边线或绘图框仍须渲染复核。",
            ],
        }


def enabled_attr(e, name): return val(e, name, "0") not in {"0", "false", "off"}
def bool_prop(pp, name): return pvalue(pp, name, default="0") not in {"0", "false", "off"}
def md_value(value):
    if isinstance(value, str): return value.replace("\n", " ")
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def write_markdown(report, path):
    s = report["summary"]
    out = ["# 论文格式审核报告", "", f"文件：`{report['source']}`", "",
           f"SHA-256：`{report['source_sha256']}`。只读检查，源文件核验未变。", "",
           f"检查 {s['paragraph_count']} 个段落、{s['table_count']} 个表格、{s['section_count']} 个节，"
           f"识别 {s['classified_paragraph_count']} 个段落角色，记录 {s['issue_count']} 项。", "",
           "本报告是静态格式审核。没有原生 Word/PDF 逐页检查，不能作最终合格结论。", "",
           "涉及字符或域显示的项目应核对已有授权及转换记录；已有授权应继续执行，本分类不要求重复询问。", ""]
    grouped = defaultdict(list)
    for issue in report["issues"]: grouped[issue["category"]].append(issue)
    priority = {"front.school_cover_missing_or_unconfirmed": 0, "front.pledge_missing": 1,
                "header.odd_even_disabled": 2, "header.wrong_effective_text": 3,
                "heading.not_true_automatic_numbering": 4, "math.table_layout_disallowed": 5,
                "math.textbox_layout_disallowed": 6, "structure.acknowledgement_not_last": 7,
                "structure.acknowledgement_missing_or_unconfirmed": 7, "heading.automatic_number_format": 8,
                "math.label_incomplete_or_missing": 9, "math.static_number": 10,
                "math.plain_text_candidate": 10, "pagination.no_page_field": 11,
                "toc.no_real_field": 12, "toc.cached_number_prefix_mismatch": 12,
                "heading.style_numbering_binding_unconfirmed": 13}
    severity_rank = {"error": 0, "warning": 1, "info": 2}
    categories = sorted(grouped, key=lambda cat: (priority.get(cat, 100), min(severity_rank[i["severity"]] for i in grouped[cat]), cat))
    main_categories = [cat for cat in categories if any(i["severity"] in {"error", "warning"} for i in grouped[cat])][:8]
    out += ["## 首要问题", "", "| 项目 | 数量 | 典型定位 |", "|---|---:|---|"]
    for cat in main_categories:
        sample = grouped[cat][0]
        loc = sample["location"]
        label = " / ".join(str(loc[k]) for k in ("section_id", "table_id", "paragraph_id", "region") if k in loc) or loc.get("part", "文档")
        out.append(f"| {cat} | {len(grouped[cat])} | {label} |")
    if not main_categories: out.append("| 未检出确定的静态错误/警告 | 0 | 仍须渲染核查 |")
    out += ["", "## 分类", "", "| 类别 | 数量 | 处理方式 |", "|---|---:|---|"]
    for status, label in FORMAT_STATUSES.items(): out.append(f"| {status} | {s['by_status'].get(status, 0)} | {label} |")
    out += ["", "## 按规则归组的问题", "", "同一规则展示最多 5 个典型定位；JSON 保留所有问题的完整定位、实际证据及建议。", ""]
    for cat in categories:
        items = grouped[cat]
        representative = items[0]
        counts = Counter(i["severity"] for i in items)
        conf = Counter(i["confidence"] for i in items)
        status = Counter(i["status"] for i in items)
        out += [f"### {cat} · {len(items)} 项", "",
                "严重程度：" + md_value(dict(counts)) + "；识别置信度：" + md_value(dict(conf)) + "。", "",
                "处理：" + "；".join(f"{FORMAT_STATUSES[k]} {v} 项" for k,v in status.items()) + "。", "",
                "目标：" + md_value(representative["expected"]), "", "建议：" + representative["recommendation"], "",
                "典型证据：", ""]
        for issue in items[:5]:
            loc = issue["location"]
            label = " / ".join(str(loc[k]) for k in ("part", "section_id", "table_id", "paragraph_id", "row_index", "header_type", "footer_type") if k in loc)
            snippet = "；原文片段：“" + loc["text_excerpt"] + "”" if loc.get("text_excerpt") else ""
            out.append(f"- `{issue['id']}` · `{label}`{snippet}；实际：{md_value(issue['actual'])}")
        if len(items) > 5: out.append(f"- 另 {len(items)-5} 项相同规则的问题，完整定位见 JSON 的 category={cat}。")
        out.append("")
    out += ["## 覆盖不足与不确定项", "",
            f"已识别 {s['classified_paragraph_count']}/{s['paragraph_count']} 个段落角色；空段和未识别内容不等于合格。", "",
            "- 中低置信度角色、未分类表格和数学变量/公式语义应先阅读原文确认。",
            "- TOC/SEQ/REF 缓存、实际页位置、奇偶页眉及图片字体须接续渲染复查。", ""]
    out += ["## 检测限制", ""] + ["- " + x for x in report["limitations"]] + [""]
    Path(path).write_text("\n".join(out), encoding="utf-8")


def main():
    ap = argparse.ArgumentParser(description="只读审核福州理工学院经管论文 DOCX 格式，输出 JSON 与中文 Markdown。")
    ap.add_argument("source", type=Path)
    ap.add_argument("--mapping", type=Path, help="可选、SHA-256 绑定的 role_map.json")
    ap.add_argument("--title", help="已确认的本论文完整题目，用于页眉比对")
    ap.add_argument("--json", dest="json_path", type=Path, required=True)
    ap.add_argument("--markdown", dest="md_path", type=Path, required=True)
    args = ap.parse_args()
    try:
        outputs = [args.json_path.resolve(), args.md_path.resolve()]
        if len(set(outputs)) != 2 or args.source.resolve() in outputs:
            raise ValueError("报告路径必须互不相同且不能与输入论文相同。")
        for path in outputs:
            if path.exists(): raise ValueError("不覆盖已有报告，请用新路径：" + str(path))
            path.parent.mkdir(parents=True, exist_ok=True)
        mapping = json.loads(args.mapping.read_text(encoding="utf-8-sig")) if args.mapping else None
        report = Audit(args.source, mapping, args.title).run()
        args.json_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        write_markdown(report, args.md_path)
        print(json.dumps({"ok": True, "source_unchanged": True, "json": str(args.json_path),
                          "markdown": str(args.md_path), "summary": report["summary"]}, ensure_ascii=False))
        return 0
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
