#!/usr/bin/env python3
"""Vendor-neutral JSON adapter for the existing, content-preserving CLI tools.

Python 3.10+. Run ``python scripts/agent_cli.py --request request.json`` or
``--request -`` for UTF-8 JSON on stdin. Relative paths use the caller's working
directory. A host may expose this process using agent-tools.json; no API vendor,
network access, shell interpolation, document renderer or agent SDK is required.
"""
from __future__ import annotations

import argparse
import hashlib
import importlib.metadata
import importlib.util
import json
import os
from pathlib import Path
import platform
import shutil
import subprocess
import sys
import tempfile

SCRIPTS = Path(__file__).resolve().parent
ACTIONS = ("doctor", "inspect", "audit", "apply_styles", "verify", "apply_layout",
           "apply_citations", "validate_controls", "prepare", "delivery_check")
SCOPES = {"styles": "thesis_format.py", "layout": "thesis_layout.py",
          "citations": "citation_style.py"}
COMMON_LIMITATIONS = [
    "Only the requested static operation is performed; success is not a certificate of complete thesis compliance.",
    "No rendering, visual page review, native field update or displayed page-parity check is performed by this adapter.",
    "Existing-paper cover replacement, chapter renumbering, equation transcription and appendix relocation require separately reviewed transformations; this adapter does not reconstruct them automatically.",
    "Each formatting stage must use the exact preceding DOCX and its reviewed SHA-256-bound map. Keep the original and all stage reports; do not reset a baseline to hide changes.",
]
REMAINING_STEPS = [
    "Check complete task coverage and rerun the final DOCX format audit; resolve or explicitly report remaining issues.",
    "Verify the complete source-to-final preservation chain, including any separately authorized structural transformations.",
    "Open the exact final DOCX in the target document application and render it successfully; inspect every final page and editable controls.",
    "Pass a separate delivery check bound to the exact final DOCX. No basic action in this interface authorizes a completion claim.",
]


class RequestError(Exception):
    def __init__(self, code, message):
        super().__init__(message)
        self.code = code


class Parser(argparse.ArgumentParser):
    def error(self, message):
        raise RequestError("INVALID_ARGUMENTS", message)


def sha(path):
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def object_pairs(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise RequestError("INVALID_JSON", "Duplicate JSON key: " + key)
        result[key] = value
    return result


def decode_json(text):
    try:
        return json.loads(text.lstrip("\ufeff"), object_pairs_hook=object_pairs,
                          parse_constant=lambda value: (_ for _ in ()).throw(
                              RequestError("INVALID_JSON", "Invalid JSON constant: " + value)))
    except json.JSONDecodeError as exc:
        raise RequestError("INVALID_JSON", str(exc)) from exc


def file_path(value, key, must_exist):
    if not isinstance(value, str) or not value.strip() or "\x00" in value:
        raise RequestError("INVALID_REQUEST", key + " must be a nonempty local path string")
    path = Path(value).expanduser().resolve()
    if must_exist and not path.is_file():
        raise RequestError("INPUT_NOT_FOUND", key + " is not a readable file: " + str(path))
    return path


def module_available(name):
    try:
        return importlib.util.find_spec(name) is not None
    except (ImportError, ValueError, ModuleNotFoundError):
        return False


def doctor():
    dependencies = {}
    for module, distribution in (("lxml", "lxml"), ("docx", "python-docx")):
        available = module_available(module)
        try:
            version = importlib.metadata.version(distribution) if available else None
        except importlib.metadata.PackageNotFoundError:
            version = None
        dependencies[distribution] = {"available": available, "version": version}
    word_registered = False
    if os.name == "nt":
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CLSID"):
                word_registered = True
        except OSError:
            pass
    return {
        "python": {"version": platform.python_version(), "executable": sys.executable,
                   "supported": sys.version_info >= (3, 10)},
        "dependencies": dependencies,
        "static_cli_ready": sys.version_info >= (3, 10) and dependencies["lxml"]["available"],
        "example_authoring_ready": all(d["available"] for d in dependencies.values()),
        "render_candidates": {
            "libreoffice_on_path": shutil.which("libreoffice") or shutil.which("soffice"),
            "word_registered_windows": word_registered,
            "pdf_page_image_tools_on_path": {name: shutil.which(name) for name in ("pdftoppm", "mutool")},
            "operationally_tested": False,
            "note": "Discovery only. Nothing installed or launched. A host-authorized renderer and required fonts must be tested separately; pagination can differ between renderers.",
        },
        "actions": list(ACTIONS),
    }


def plan(request, request_file=None):
    if not isinstance(request, dict):
        raise RequestError("INVALID_REQUEST", "Request must be a JSON object")
    action = request.get("action")
    if not isinstance(action, str) or action not in ACTIONS:
        raise RequestError("UNKNOWN_ACTION", "Allowed actions: " + ", ".join(ACTIONS))
    allowed = {"action"}
    required = set()
    reads, writes = {}, {}
    if action == "inspect":
        allowed |= {"source", "report", "scope"}; required |= {"source", "report"}
    elif action == "audit":
        allowed |= {"source", "report", "markdown", "mapping", "title"}
        required |= {"source", "report", "markdown"}
    elif action.startswith("apply_") or action == "verify":
        allowed |= {"source", "output", "mapping", "report"}
        required |= {"source", "output", "mapping", "report"}
        if action == "verify":
            allowed.add("scope")
    elif action == "validate_controls":
        allowed |= {"source", "report"}; required |= {"source", "report"}
    extra, missing = set(request) - allowed, required - set(request)
    if extra or missing:
        raise RequestError("INVALID_REQUEST", "Unexpected fields: " + ", ".join(sorted(extra)) +
                           "; missing fields: " + ", ".join(sorted(missing)))
    scope = request.get("scope", "styles")
    if not isinstance(scope, str) or scope not in SCOPES:
        raise RequestError("INVALID_REQUEST", "scope must be styles, layout or citations")
    if "title" in request and (not isinstance(request["title"], str) or not request["title"].strip()):
        raise RequestError("INVALID_REQUEST", "title must be a nonempty string")
    for key in ("source", "mapping"):
        if key in request:
            reads[key] = file_path(request[key], key, True)
    if "output" in request:
        collection = reads if action == "verify" else writes
        collection["output"] = file_path(request["output"], "output", action == "verify")
    for key in ("report", "markdown"):
        if key in request:
            writes[key] = file_path(request[key], key, False)
    for key in ("source", "output"):
        path = reads.get(key) or writes.get(key)
        if path is not None and path.suffix.lower() != ".docx":
            raise RequestError("UNSUPPORTED_INPUT", key + " must be a DOCX file")
    if len(set(writes.values())) != len(writes):
        raise RequestError("OUTPUT_CONFLICT", "Output paths must be distinct")
    protected = set(reads.values()) | ({request_file.resolve()} if request_file else set())
    for path in writes.values():
        if path in protected:
            raise RequestError("OUTPUT_CONFLICT", "Output collides with an input: " + str(path))
        if path.exists() or path.is_symlink():
            raise RequestError("OUTPUT_EXISTS", "Refusing to overwrite: " + str(path))
    if "mapping" in reads:
        mapping = decode_json(reads["mapping"].read_text(encoding="utf-8-sig"))
        if not isinstance(mapping, dict) or mapping.get("source_sha256") != sha(reads["source"]):
            raise RequestError("STALE_MAPPING", "Mapping source_sha256 is missing or differs from this exact source; inspect and review again")
    return action, scope, reads, writes


def command_for(action, scope, request, reads, staged):
    def p(key):
        return str(staged[key] if key in staged else reads[key])
    if action == "inspect":
        return [SCOPES[scope], "inspect", p("source"), "--out", p("report")]
    if action == "audit":
        result = ["format_audit.py", p("source"), "--json", p("report"), "--markdown", p("markdown")]
        if "mapping" in reads:
            result += ["--mapping", p("mapping")]
        if "title" in request:
            result += ["--title", request["title"]]
        return result
    if action == "validate_controls":
        return ["validate_academic_controls.py", p("source"), "--report", p("report")]
    if action == "apply_citations" or (action == "verify" and scope == "citations"):
        return ["citation_style.py", "verify" if action == "verify" else "apply",
                "--source", p("source"), "--output", p("output"), "--map", p("mapping"), "--report", p("report")]
    script = SCOPES[scope] if action == "verify" else SCOPES[action.removeprefix("apply_")]
    arguments = (["verify", p("output"), "--source", p("source")] if action == "verify"
                 else ["apply", p("source"), p("output")])
    return [script, *arguments, "--map", p("mapping"), "--report", p("report")]


def replace_paths(value, replacements):
    if isinstance(value, dict):
        return {key: replace_paths(item, replacements) for key, item in value.items()}
    if isinstance(value, list):
        return [replace_paths(item, replacements) for item in value]
    if isinstance(value, str):
        for old, new in replacements.items():
            value = value.replace(old, new)
    return value


def execute_workflow(request, request_file=None):
    """Expose the preparation and delivery CLIs without bypassing their checks.

    prepare owns its new directory and records partial failures there. Delivery
    reports stage into a temporary file and publish byte-for-byte, preserving
    all fresh hashes; exit 3 is a completed check with unresolved blockers.
    """
    action=request['action']
    required={'action','source','work_dir'} if action=='prepare' else {'action','source','output','report'}
    optional=set() if action=='prepare' else {'chain','render_manifest','page_manifest','review'}
    extra=set(request)-required-optional; missing=required-set(request)
    if extra or missing:
        raise RequestError('INVALID_REQUEST','Unexpected fields: '+', '.join(sorted(extra))+'; missing fields: '+', '.join(sorted(missing)))
    reads={key:file_path(request[key],key,True) for key in ('source','output','chain','render_manifest','page_manifest','review') if key in request}
    for key in ('source','output'):
        if key in reads and reads[key].suffix.lower()!='.docx':raise RequestError('UNSUPPORTED_INPUT',key+' must be a DOCX file')
    write_key='work_dir' if action=='prepare' else 'report'
    destination=file_path(request[write_key],write_key,False)
    protected=set(reads.values())|({request_file.resolve()} if request_file else set())
    if destination in protected:raise RequestError('OUTPUT_CONFLICT','Output collides with an input: '+str(destination))
    if destination.exists() or destination.is_symlink():raise RequestError('OUTPUT_EXISTS','Refusing to overwrite: '+str(destination))
    before={key:sha(path) for key,path in reads.items()}
    response={'protocol_version':1,'action':action,'ok':False,'operation_succeeded':False,
              'delivery_ready':False,'status':'failed','outputs':{},'remaining_steps':list(REMAINING_STEPS),
              'render_review_performed':False,'checks':{'static_operation':'not_run','visual_review':'not_performed_by_this_adapter'}}
    with tempfile.TemporaryDirectory(prefix='fit-workflow-') as temporary:
        staged=Path(temporary)/'delivery-report.json'
        if action=='prepare':
            arguments=['prepare.py','--source',str(reads['source']),'--work-dir',str(destination)]
        else:
            arguments=['delivery_check.py','--source',str(reads['source']),'--final',str(reads['output']),'--report',str(staged)]
            for key in optional:
                if key in reads:arguments+=['--'+key.replace('_','-'),str(reads[key])]
        completed=subprocess.run([sys.executable,str(SCRIPTS/arguments[0]),*arguments[1:]],shell=False,
                                 capture_output=True,text=True,encoding='utf-8',errors='replace',
                                 env=dict(os.environ,PYTHONUTF8='1',PYTHONIOENCODING='utf-8'))
        unchanged={key:path.is_file() and sha(path)==before[key] for key,path in reads.items()}
        response['checks']['input_files_unchanged']=unchanged
        response['source_sha256']=before['source']; response['tool_exit_code']=completed.returncode
        if action=='prepare' and destination.is_dir():
            response['outputs']['work_dir']=str(destination)
            if (destination/'prepare-result.json').is_file():response['outputs']['report']=str(destination/'prepare-result.json')
        if not all(unchanged.values()):
            response['error']={'code':'INPUT_CHANGED','message':'An input changed during execution; result is not accepted'}
            return response
        try:
            result=decode_json(completed.stdout.strip())
        except RequestError:
            result={'message':(completed.stdout or completed.stderr)[-8000:]}
        response['result']=result
        protocol_valid=(isinstance(result,dict) and type(result.get('operation_succeeded')) is bool
                        and type(result.get('delivery_ready')) is bool)
        if not protocol_valid:
            response['error']={'code':'TOOL_PROTOCOL_ERROR','message':'Workflow must return explicit boolean operation_succeeded and delivery_ready'}
            return response
        success=(result['operation_succeeded'] is True and
                 ((action=='prepare' and completed.returncode==0 and result['delivery_ready'] is False) or
                  (action=='delivery_check' and completed.returncode==(0 if result['delivery_ready'] else 3))))
        if not success:
            response['status']=result.get('status','failed')
            response['remaining_steps']=result.get('remaining_steps',response['remaining_steps'])
            response['error']={'code':'TOOL_FAILED','message':str(result.get('error','Workflow execution failed or returned inconsistent status'))}
            return response
        if action=='delivery_check':
            if not staged.is_file():
                response['error']={'code':'MISSING_OUTPUT','message':'Delivery checker did not write its fresh report'}
                return response
            destination.parent.mkdir(parents=True,exist_ok=True)
            try:
                with destination.open('xb') as handle:handle.write(staged.read_bytes())
            except FileExistsError:
                response['error']={'code':'OUTPUT_EXISTS','message':'Delivery report appeared during execution; not overwritten'}
                return response
            response['outputs']['report']=str(destination)
            result=dict(result,report=str(destination));response['result']=result
            response['blocking_items']=result.get('blocking_items',[])
        elif not (destination/'prepare-result.json').is_file():
            response['error']={'code':'MISSING_OUTPUT','message':'Preparation did not write its completion state'}
            return response
        response.update(ok=True,operation_succeeded=True,delivery_ready=result['delivery_ready'],status=result.get('status','operation_completed'))
        response['checks']['static_operation']='completed'
        response['remaining_steps']=([] if result['delivery_ready'] else
            result.get('remaining_steps',['Resolve blocking_items and rerun delivery_check on the exact final files'] if action=='delivery_check' else list(REMAINING_STEPS)))
        response['limitations']=['This adapter did not perform visual inspection; delivery_check validates the supplied page-review evidence and fresh checks.']
        return response


def execute(request, request_file=None):
    if isinstance(request,dict) and request.get('action') in ('prepare','delivery_check'):
        return execute_workflow(request,request_file)
    action, scope, reads, writes = plan(request, request_file)
    response = {"protocol_version": 1, "action": action, "ok": False,
                "operation_succeeded": False, "delivery_ready": False, "status": "failed",
                "remaining_steps": [] if action == "doctor" else list(REMAINING_STEPS),
                "render_review_performed": False, "needs_render_review": action != "doctor",
                "outputs": {}, "checks": {"static_operation": "not_run", "rendering": "not_run",
                                           "visual_review": "not_run", "native_field_update": "not_run"},
                "limitations": list(COMMON_LIMITATIONS)}
    if action == "doctor":
        response.update(ok=True, operation_succeeded=True, status="capabilities_checked", result=doctor())
        return response
    if not module_available("lxml"):
        raise RequestError("DEPENDENCY_MISSING", "lxml is required; run doctor and install requirements in the host-approved Python environment")
    snapshots = {key: sha(path) for key, path in reads.items()}
    with tempfile.TemporaryDirectory(prefix="fit-agent-") as temporary:
        staged = {key: Path(temporary) / (key + (".docx" if key == "output" else ".md" if key == "markdown" else ".json")) for key in writes}
        args = command_for(action, scope, request, reads, staged)
        argv = [sys.executable, str(SCRIPTS / args[0]), *args[1:]]
        env = dict(os.environ, PYTHONUTF8="1", PYTHONIOENCODING="utf-8")
        completed = subprocess.run(argv, shell=False, capture_output=True, text=True, encoding="utf-8", errors="replace", env=env)
        unchanged = {key: path.is_file() and sha(path) == snapshots[key] for key, path in reads.items()}
        response["checks"].update(static_operation="completed" if completed.returncode == 0 else "failed",
                                  input_files_unchanged=unchanged)
        response["source_sha256"] = snapshots["source"]
        if not all(unchanged.values()):
            response["error"] = {"code": "INPUT_CHANGED", "message": "An input changed during execution; no staged result published"}
            return response
        replacements = {str(staged[key]): str(path) for key, path in writes.items()}
        raw = completed.stdout.strip() or completed.stderr.strip()
        protocol_valid=True
        try:
            result = decode_json(raw) if raw else {}
        except RequestError:
            protocol_valid=False
            result = {"message": raw[-8000:]}
        result = replace_paths(result, replacements)
        response.update(result=result, tool_exit_code=completed.returncode)
        success_key='passed' if action=='validate_controls' else 'ok'
        protocol_valid=protocol_valid and isinstance(result,dict) and type(result.get(success_key)) is bool
        operation_ok = completed.returncode == 0 and protocol_valid and result[success_key] is True
        response['checks']['static_operation']='completed' if operation_ok else 'failed'
        # Failed verification still produces a useful report, but a failed apply
        # must never publish a document as if it were an accepted result.
        publish = {key: path for key, path in staged.items() if path.is_file() and (key != "output" or operation_ok)}
        for key, path in publish.items():
            destination = writes[key]
            destination.parent.mkdir(parents=True, exist_ok=True)
            data = path.read_bytes()
            if key == "report":
                try:
                    data = json.dumps(replace_paths(decode_json(data.decode("utf-8-sig")), replacements), ensure_ascii=False, indent=2).encode("utf-8")
                except (RequestError, UnicodeDecodeError):
                    pass
            elif key == "markdown":
                data = replace_paths(data.decode("utf-8"), replacements).encode("utf-8")
            try:
                with destination.open("xb") as handle:
                    handle.write(data)
            except FileExistsError:
                response["error"] = {"code": "OUTPUT_EXISTS", "message": "Output appeared during execution; not overwritten: " + str(destination)}
                return response
            response["outputs"][key] = str(destination)
        if operation_ok and set(response["outputs"]) != set(writes):
            response["error"] = {"code": "MISSING_OUTPUT", "message": "Underlying tool reported success without all requested outputs"}
        elif not protocol_valid:
            response["error"] = {"code": "TOOL_PROTOCOL_ERROR", "message": "Underlying tool must emit one JSON object with an explicit boolean " + success_key + "; no DOCX published"}
        elif not operation_ok:
            response["error"] = {"code": "CHECK_FAILED" if action in {"verify", "validate_controls"} else "TOOL_FAILED",
                                 "message": "Underlying tool did not pass; see result and any published diagnostic report"}
        else:
            response.update(ok=True, operation_succeeded=True, status="operation_completed_not_delivery")
        if action == "audit":
            response["limitations"].append("Audit findings are heuristics and static checks. ok=true means reports were produced, not that the thesis passed every rule.")
        if action == "validate_controls":
            response["limitations"].append("validate_controls checks this package's native numbering/OMML convention for new examples; it is not a generic verdict on every existing thesis.")
        return response


def main():
    for stream in (sys.stdin, sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
    action = None
    try:
        parser = Parser(description=__doc__)
        parser.add_argument("--request", required=True, help="UTF-8 JSON file, or - for stdin")
        args = parser.parse_args()
        request_file = None if args.request == "-" else file_path(args.request, "request", True)
        raw = sys.stdin.read() if request_file is None else request_file.read_text(encoding="utf-8-sig")
        request = decode_json(raw)
        action = request.get("action") if isinstance(request, dict) else None
        result = execute(request, request_file)
    except RequestError as exc:
        result = {"protocol_version": 1, "action": action, "ok": False, "operation_succeeded": False,
                  "delivery_ready": False, "status": "failed", "remaining_steps": list(REMAINING_STEPS),
                  "outputs": {}, "render_review_performed": False,
                  "error": {"code": exc.code, "message": str(exc)},
                  "checks": {"static_operation": "not_run", "rendering": "not_run", "visual_review": "not_run", "native_field_update": "not_run"}}
    except (OSError, UnicodeError, ValueError) as exc:
        result = {"protocol_version": 1, "action": action, "ok": False, "operation_succeeded": False,
                  "delivery_ready": False, "status": "failed", "remaining_steps": list(REMAINING_STEPS),
                  "outputs": {}, "render_review_performed": False,
                  "error": {"code": "IO_ERROR", "message": str(exc)},
                  "checks": {"rendering": "not_run", "visual_review": "not_run", "native_field_update": "not_run"}}
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result["ok"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
