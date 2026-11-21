"""Interactive sessions for the Excel console. Only JSON previews cross the bridge."""
import ast
import builtins
import codeop
from collections import OrderedDict
from itertools import chain, islice
import keyword
import reprlib
import sys
import threading
import traceback
import types


PAGE_SIZE = 50
HANDLE_LIMIT = 2048
SCALARS = (str, bytes, int, float, bool, complex, type(None))
BRACKETS = {dict: ("{", "}"), list: ("[", "]"), tuple: ("(", ")"),
            set: ("{", "}"), frozenset: ("frozenset({", "})")}


def _brief(value):
    if type(value) in SCALARS:
        return reprlib.repr(value)
    for base, (left, right) in BRACKETS.items():
        if isinstance(value, base):
            return left + ("…" if base.__len__(value) else "") + right
    return type(value).__name__


def _collection_preview(value, base):
    count = base.__len__(value)
    if base is dict:
        items = [f"{_brief(k)}: {_brief(v)}" for k, v in islice(dict.items(value), 3)]
    else:
        items = [_brief(item) for item in islice(base.__iter__(value), 3)]
    if count > len(items):
        items.append("…")
    left, right = BRACKETS[base]
    contents = ", ".join(items) + ("," if base is tuple and count == 1 else "")
    return f"{type(value).__name__} ({count}) {left}{contents}{right}"


def _namespaces(obj):
    """Read stored fields and class dictionaries without calling user properties."""
    cls = type(obj)
    if not isinstance(obj, type):
        descriptor = next((vars(base)["__dict__"] for base in cls.__mro__
                           if "__dict__" in vars(base)), None)
        if type(descriptor) in (types.GetSetDescriptorType, types.MemberDescriptorType):
            fields = descriptor.__get__(obj, cls)
            if isinstance(fields, dict):
                yield fields
    for base in (obj if isinstance(obj, type) else cls).__mro__:
        yield vars(base)


def _attribute(obj, name):
    for fields in _namespaces(obj):
        if name in fields:
            value = fields[name]
            if type(value) is types.MemberDescriptorType and not isinstance(obj, type):
                return value.__get__(obj)
            return value
    raise AttributeError(name)


def _fields(obj):
    seen = set()
    for fields in _namespaces(obj):
        for name, value in fields.items():
            if not isinstance(name, str) or name.startswith("__") or name in seen:
                continue
            seen.add(name)
            if type(value) is types.MemberDescriptorType and not isinstance(obj, type):
                try:
                    value = value.__get__(obj)
                except AttributeError:
                    continue
            yield name, value


def _items(obj):
    if isinstance(obj, dict):
        return ((reprlib.repr(k), v) for k, v in dict.items(obj))
    for base in (list, tuple, set, frozenset, range):
        if isinstance(obj, base):
            return ((str(i), v) for i, v in enumerate(base.__iter__(obj)))
    return _fields(obj)


class _Stream:
    def __init__(self, console, kind):
        self.console = console
        self.kind = kind

    def write(self, text):
        if not isinstance(text, str):
            raise TypeError("write() requires a string")
        if text:
            self.console.emit(self.kind, text)
        return len(text)

    def flush(self):
        pass


class Console:
    def __init__(self, capture, session):
        self.capture = capture
        self.console_stream = sys.stdout
        self.stdout = _Stream(self, "stdout")
        self.stderr = _Stream(self, "stderr")
        self.module_name = "_excel_console_" + session
        module = types.ModuleType(self.module_name)
        self.namespace = module.__dict__
        self.namespace["__builtins__"] = dict(vars(builtins), print=self.print)
        sys.modules[self.module_name] = module
        self.compiler = codeop.CommandCompiler()
        self.objects = OrderedDict()
        self.next_handle = 0
        self.events = []
        self.event_lock = threading.Lock()
        self.output_size = 0
        self.output_count = 0

    def close(self):
        sys.modules.pop(self.module_name, None)
        self.objects.clear()

    def describe(self, value, *, printed=False):
        if type(value) in SCALARS:
            return {"text": reprlib.repr(value)}
        name = type(value).__name__
        for container in BRACKETS:
            if isinstance(value, container):
                name = _collection_preview(value, container)
                break
        else:
            if type(value).__repr__ is not object.__repr__ or (printed and type(value).__str__ is not object.__str__):
                name = str(value)[:240] if printed else reprlib.repr(value)
        self.next_handle += 1
        handle = self.next_handle
        self.objects[handle] = value
        if len(self.objects) > HANDLE_LIMIT:
            self.objects.popitem(last=False)
        return {"text": name, "handle": handle}

    def children(self, handle, offset):
        if handle not in self.objects:
            raise ValueError("This object preview has expired. Evaluate the object again to inspect it.")
        self.objects.move_to_end(handle)
        page = list(islice(_items(self.objects[handle]), offset, offset + PAGE_SIZE + 1))
        return {"items": [dict(label=k, **self.describe(v)) for k, v in page[:PAGE_SIZE]],
                "more": len(page) > PAGE_SIZE}

    def emit(self, kind, value):
        with self.event_lock:
            if self.output_count > 2000 or self.output_size > 262144:
                return
            size = len(value) if isinstance(value, str) else 1
            remaining = 262144 - self.output_size
            if size <= remaining or remaining and isinstance(value, str):
                self.events.append({"kind": kind, "value": value[:remaining] if isinstance(value, str) else value})
            self.output_size += size
            self.output_count += 1
            if self.output_count > 2000 or self.output_size > 262144:
                self.events.append({"kind": "stderr", "value": "\n[Console output truncated]\n"})

    def print(self, *values, sep=" ", end="\n", file=None, flush=False):
        if ((file if file is not None else sys.stdout) is not self.console_stream
                or getattr(self.capture, "stdout", None) is not self.stdout):
            return builtins.print(*values, sep=sep, end=end, file=file, flush=flush)
        sep = " " if sep is None else sep
        end = "\n" if end is None else end
        if not isinstance(sep, str) or not isinstance(end, str):
            raise TypeError("sep and end must be strings or None")
        for i, value in enumerate(values):
            if i:
                self.stdout.write(sep)
            if type(value) in SCALARS:
                self.stdout.write(str(value))
            else:
                self.emit("object", self.describe(value, printed=True))
        self.stdout.write(end)
        if flush:
            self.stdout.flush()

    def complete(self, token):
        parts = token.split(".")
        prefix = parts.pop()
        if (prefix and not prefix.isidentifier()) or any(not part.isidentifier() for part in parts):
            return []
        if not parts:
            names = chain(self.namespace, self.namespace["__builtins__"], keyword.kwlist)
        else:
            root = parts.pop(0)
            try:
                obj = self.namespace[root] if root in self.namespace else vars(builtins)[root]
                for part in parts:
                    obj = _attribute(obj, part)
            except (AttributeError, KeyError):
                return []
            names = chain.from_iterable(_namespaces(obj))
        return sorted({n for n in names if isinstance(n, str) and n.startswith(prefix)
                       and (prefix.startswith("_") or not n.startswith("_"))})[:80]

    def run(self, source):
        self.output_size = self.output_count = 0
        previous = self.capture.__dict__.copy()
        self.capture.stdout, self.capture.stderr = self.stdout, self.stderr
        filename = "<Python console>"
        try:
            code = self.compiler(source, filename, "exec")
            if code is None:
                return {"incomplete": True}
            tree = compile(source, filename, "exec", code.co_flags | ast.PyCF_ONLY_AST, dont_inherit=True)
            if tree.body and isinstance(tree.body[-1], ast.Expr):
                last = tree.body.pop().value
                exec(compile(tree, filename, "exec", code.co_flags, dont_inherit=True), self.namespace)
                result = eval(compile(ast.Expression(last), filename, "eval", code.co_flags, dont_inherit=True),
                              self.namespace)
                if result is not None:
                    self.namespace["_"] = result
                    self.emit("result", self.describe(result))
            else:
                exec(code, self.namespace)
        except BaseException as exc:
            tb = exc.__traceback__
            while tb and tb.tb_frame.f_code.co_filename != filename:
                tb = tb.tb_next
            self.stderr.write("".join(traceback.format_exception(type(exc), exc, tb)))
        finally:
            self.capture.__dict__.clear()
            self.capture.__dict__.update(previous)
        return {"incomplete": False}

    def request(self, action, args):
        if action == "run":
            return self.run(args["source"])
        if action == "drain":
            with self.event_lock:
                events, self.events = self.events, []
                return events
        if action == "complete":
            return self.complete(args["token"])
        if action == "children":
            return self.children(args["handle"], args["offset"])
        if action == "clear":
            self.objects.clear()
            return None
        raise ValueError(f"Unknown console action: {action}")
