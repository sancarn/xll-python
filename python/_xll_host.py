"""In-process CPython entry point used by the native XLL bridge."""
import asyncio
import dataclasses
import datetime
import hashlib
import importlib
import importlib.machinery
import importlib.util
import inspect
import io
import json
import linecache
import math
import os
from pathlib import Path
import re
import site
import sys
import threading
import traceback
import types
import typing

from xll_python import ExcelError, Function, Reference
from _xll_repl import Console


_functions = {}
_managed_paths = []
_script_lock = threading.RLock()
_capture = threading.local()
_debug_lock = threading.Lock()
_debug = None


class _Output:
    encoding = "utf-8"

    def __init__(self, original, channel):
        self.original = original
        self.channel = channel

    def destination(self):
        return getattr(_capture, self.channel, None) or getattr(_capture, "stream", None) or self.original

    def write(self, data):
        stream = self.destination()
        return stream.write(data) if stream is not None else len(data)

    def flush(self):
        stream = self.destination()
        if stream is not None:
            stream.flush()

    def isatty(self):
        return False

    def __getattr__(self, name):
        if self.original is not None:
            return getattr(self.original, name)
        raise AttributeError(name)


sys.stdout = _Output(sys.stdout, "stdout")
sys.stderr = _Output(sys.stderr, "stderr")
_consoles = {}


def _annotation(annotation):
    if typing.get_origin(annotation) is typing.Annotated:
        return typing.get_args(annotation)
    return (annotation,)


def _cell_name(name):
    if re.fullmatch(r"R(?:[1-9][0-9]*)?C(?:[1-9][0-9]*)?|R|C", name, re.I):
        return True
    match = re.fullmatch(r"([A-Z]{1,3})([1-9][0-9]{0,6})", name.upper())
    if match is None:
        return False
    column = 0
    for letter in match[1]:
        column = 26 * column + ord(letter) - ord("A") + 1
    return column <= 16384 and int(match[2]) <= 1048576


def _register(fn, functions, entries):
    opt = getattr(fn, "__excel_function__", None)
    raw_return = fn.__annotations__.get("return")
    if opt is None and not (typing.get_origin(raw_return) is typing.Annotated or
                            isinstance(raw_return, str) and "Annotated" in raw_return):
        return
    hints = typing.get_type_hints(fn, include_extras=True)
    ret = _annotation(hints.get("return"))
    opt = opt or next((x for x in ret[1:] if isinstance(x, Function)), None)
    if opt is None:
        return
    name = opt.name or fn.__name__
    if (not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]{0,254}", name)
        or _cell_name(name)):
        raise ValueError(f"Invalid or cell-like Excel function name: {name}")
    if name.upper() in functions:
        raise ValueError(f"Duplicate worksheet function: {name}")
    signature = inspect.signature(fn)
    args = []
    for par in signature.parameters.values():
        if par.kind not in (par.POSITIONAL_ONLY, par.POSITIONAL_OR_KEYWORD):
            raise TypeError(f"{name} needs a fixed positional signature")
        meta = _annotation(hints.get(par.name))[1:]
        args.append(dict(name=par.name, reference=_annotation(hints.get(par.name))[0] is Reference,
                         description=next((m for m in meta if isinstance(m, str)), "")))
    asynchronous = opt.asynchronous if opt.asynchronous is not None else inspect.iscoroutinefunction(fn)
    if inspect.iscoroutinefunction(fn) and not asynchronous:
        raise ValueError(f"{name}: async def requires asynchronous registration")
    if opt.macro_type and (opt.thread_safe or asynchronous):
        raise ValueError(f"{name}: incompatible Excel macro-type registration flags")
    if len(args) > (254 if asynchronous else 255):
        raise ValueError(f"{name}: too many Excel arguments")
    entries.append(dict(dataclasses.asdict(opt), name=name, arguments=args,
                        asynchronous=asynchronous, description=opt.description or inspect.getdoc(fn) or ""))
    functions[name.upper()] = (fn, signature, hints)


class _FreshSource(importlib.machinery.SourceFileLoader):
    def get_code(self, fullname):
        return self.source_to_code(self.get_data(self.path), self.path)


class _StartupImports:
    """Read local helpers from source even for same-size edits within one timestamp tick."""
    roots = ()

    def owns(self, filename):
        return isinstance(filename, str) and os.path.normcase(os.path.abspath(filename)).startswith(self.roots)

    def find_spec(self, fullname, path=None, target=None):
        if not self.roots:
            return None
        spec = importlib.machinery.PathFinder.find_spec(fullname, path)
        if spec and isinstance(spec.loader, importlib.machinery.SourceFileLoader) and self.owns(spec.origin):
            spec.loader = _FreshSource(fullname, spec.origin)
        return spec

    def modules(self):
        return {name: module for name, module in sys.modules.copy().items()
                if name not in (__name__, "xll_python", "__main__")
                and self.owns(getattr(module, "__file__", None)) and str(module.__file__).endswith(".py")}


_startup_imports = _StartupImports()
sys.meta_path.insert(sys.meta_path.index(importlib.machinery.PathFinder), _startup_imports)


def _discover(folders, packages, refresh=False):
    global _functions, _managed_paths
    previous_roots = _startup_imports.roots
    roots = tuple(os.path.join(os.path.normcase(os.path.abspath(folder)), "") for folder in folders)
    previous_modules = {}
    if refresh:
        _startup_imports.roots = previous_roots + roots
        previous_modules = _startup_imports.modules()
        for name in previous_modules:
            del sys.modules[name]
        linecache.clearcache()
    _startup_imports.roots = roots
    old_path = sys.path.copy()
    sys.path[:] = [path for path in sys.path if path not in _managed_paths]
    original = sys.path.copy()
    functions, entries, errors = {}, [], []
    try:
        for folder in packages:
            if not Path(folder).is_dir():
                errors.append(f"Missing package folder: {folder}")
            else:
                site.addsitedir(folder)
        for folder in reversed(folders):
            if folder not in sys.path:
                sys.path.insert(1, folder)
        importlib.invalidate_caches()
        for folder in map(Path, folders):
            if not folder.is_dir():
                errors.append(f"Missing function folder: {folder}")
                continue
            for path in sorted(folder.glob("*.py")):
                if path.name.startswith("_"):
                    continue
                key = "_excel_" + hashlib.sha256(os.fsencode(path)).hexdigest()
                try:
                    spec = importlib.util.spec_from_file_location(key, path)
                    module = importlib.util.module_from_spec(spec)
                    sys.modules[key] = module
                    exec(compile(path.read_bytes(), str(path), "exec"), module.__dict__)
                    local = [fn for fn in vars(module).values()
                             if inspect.isfunction(fn) and fn.__module__ == key]
                    for fn in dict.fromkeys([*local, *getattr(module, "__excel_exports__", ())]):
                        try:
                            _register(fn, functions, entries)
                        except Exception:
                            errors.append(f"{path} ({fn.__name__})\n{traceback.format_exc()}")
                except Exception:
                    errors.append(f"{path}\n{traceback.format_exc()}")
        if refresh and errors:
            raise ValueError("Refresh failed. Previous functions remain active.\n" + "\n".join(errors))
        _functions = functions
        _managed_paths = [path for path in sys.path if path not in original]
        return dict(functions=entries, errors=errors)
    except BaseException:
        sys.path[:] = old_path
        if refresh:
            for name in _startup_imports.modules():
                sys.modules.pop(name, None)
            sys.modules.update(previous_modules)
        _startup_imports.roots = previous_roots
        raise


def _decode(value):
    if isinstance(value, list):
        return [_decode(x) for x in value]
    if isinstance(value, dict) and "error" in value:
        return ExcelError(value["error"])
    if isinstance(value, dict) and "areas" in value:
        return Reference(value["sheet_id"], tuple(tuple(a) for a in value["areas"]))
    return value


def _convert(value, annotation):
    typ = _annotation(annotation)[0]
    if value is None or isinstance(value, ExcelError):
        return value
    if typing.get_origin(typ) is list:
        arg = typing.get_args(typ)[0]
        if not isinstance(value, list):
            value = [[value]] if typing.get_origin(arg) is list else [value]
        return [_convert(x, arg) for x in value]
    if typ in (int, float, bool, str):
        return typ(value)
    if typ in (datetime.datetime, datetime.date):
        dt = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=value)
        return dt if typ is datetime.datetime else dt.date()
    return value


def _encode(value):
    if isinstance(value, ExcelError):
        return {"error": value.code}
    if isinstance(value, (datetime.date, datetime.datetime)):
        if not isinstance(value, datetime.datetime):
            value = datetime.datetime.combine(value, datetime.time())
        return (value - datetime.datetime(1899, 12, 30)).total_seconds() / 86400
    if isinstance(value, (list, tuple)):
        return [_encode(x) for x in value]
    if isinstance(value, float) and not math.isfinite(value):
        return {"error": 36}
    if hasattr(value, "tolist"):
        return _encode(value.tolist())
    return value


def _invoke(registration, args):
    fn, signature, hints = registration
    values = []
    for par, value in zip(signature.parameters.values(), args):
        if isinstance(value, dict) and value.get("missing") is True:
            if par.default is par.empty:
                raise TypeError(f"Missing argument: {par.name}")
            values.append(par.default)
        else:
            values.append(_convert(_decode(value), hints.get(par.name)))
    result = fn(*values)
    if inspect.isawaitable(result):
        result = asyncio.run(result)
    return _encode(result)


def _script(path):
    path = Path(path).resolve(strict=True)
    # Serialize mutations to argv and the script search path between script runs.
    with _script_lock:
        stream = io.StringIO()
        previous = getattr(_capture, "stream", None)
        _capture.stream = stream
        argv = sys.argv
        class ScriptPath(str):
            pass
        directory = ScriptPath(str(path.parent))
        sys.path.insert(0, directory)
        sys.argv = [str(path)]
        try:
            module = types.ModuleType("__main__")
            module.__file__ = str(path)
            module.__cached__ = None
            old_main = sys.modules.get("__main__")
            sys.modules["__main__"] = module
            try:
                exec(compile(path.read_bytes(), str(path), "exec"), module.__dict__)
            finally:
                if old_main is None:
                    sys.modules.pop("__main__", None)
                else:
                    sys.modules["__main__"] = old_main
        except SystemExit as err:
            if err.code not in (None, 0):
                raise RuntimeError(f"{stream.getvalue()}Script exited with {err.code}") from err
        except BaseException as err:
            raise RuntimeError(stream.getvalue() + traceback.format_exc()) from err
        finally:
            sys.argv = argv
            sys.path[:] = [item for item in sys.path if item is not directory]
            _capture.stream = previous
        return stream.getvalue()


class _Debugger:
    def __init__(self, path, line):
        self.path = os.path.normcase(os.path.abspath(path))
        self.line = line
        self.step = False
        self.condition = threading.Condition()
        self.state = dict(status="armed")
        self.stopped = False

    def trace(self, frame, event, arg):
        if self.stopped or os.path.normcase(frame.f_code.co_filename) != self.path:
            return None
        if event == "line" and (self.step or self.line == 0 or frame.f_lineno == self.line):
            with self.condition:
                if self.stopped:
                    return None
                self.state = dict(status="paused", line=frame.f_lineno, function=frame.f_code.co_name,
                                  source=linecache.getline(frame.f_code.co_filename, frame.f_lineno).strip(),
                                  locals={k: _preview(v) for k, v in frame.f_locals.items() if not k.startswith("__")})
                self.condition.wait_for(lambda: self.state["status"] != "paused")
        return self.trace

    def command(self, command):
        with self.condition:
            if command == "stop":
                self.stopped = True
            elif self.state["status"] != "paused":
                return
            self.step = command == "step"
            if self.line == 0:
                self.line = -1
            self.state = dict(status="detached" if self.stopped else "running")
            self.condition.notify_all()


def _preview(value):
    import reprlib
    try:
        return reprlib.repr(value)
    except Exception:
        return "<unavailable>"


def _execute(operation, request):
    registration = None
    if operation == "call":
        registration = _functions.get(request["name"].upper())
        if registration is None:
            return _encode(ExcelError(29))  # Excel can still call a deleted function from an existing formula.
    with _debug_lock:
        debug = _debug
        path = request["path"] if operation == "script" else registration[0].__code__.co_filename
        if debug and debug.path == os.path.normcase(os.path.abspath(path)) and debug.state["status"] == "armed":
            debug.state = dict(status="running")
        else:
            debug = None
    previous = sys.gettrace()
    try:
        if debug:
            sys.settrace(debug.trace)
        return _script(path) if operation == "script" else _invoke(registration, request["args"])
    finally:
        if debug:
            sys.settrace(previous)
            with debug.condition:
                debug.state = dict(status="finished")


def dispatch(message):
    global _debug
    try:
        req = json.loads(message)
        op = req["op"]
        if op == "repl":
            session = req["session"]
            action = req["action"]
            if action == "open":
                _consoles[session] = Console(_capture, session)
                value = None
            elif action == "close":
                console = _consoles.pop(session, None)
                if console:
                    console.close()
                value = None
            else:
                value = _consoles[session].request(action, req.get("args", {}))
        elif op == "discover":
            value = _discover(req["folders"], req["packages"], req.get("refresh", False))
        elif op in ("call", "script"):
            value = _execute(op, req)
        elif op == "arm":
            with _debug_lock:
                if _debug:
                    _debug.command("stop")
                linecache.clearcache()
                _debug = _Debugger(req["path"], req["line"])
            value = None
        elif op == "debug":
            with _debug_lock:
                if _debug and "command" in req:
                    _debug.command(req["command"])
                value = _debug.state if _debug else dict(status="idle")
        else:
            raise ValueError(f"Unknown operation: {op}")
        return json.dumps(dict(value=value), allow_nan=False)
    except BaseException:
        return json.dumps(dict(failure=traceback.format_exc()))

