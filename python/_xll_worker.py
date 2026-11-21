"""Load one published project into a private interpreter and dispatch Excel requests."""
import importlib
import inspect
import json
import os
import shutil
import tempfile
from pathlib import Path
import sys
import threading
import traceback


def load(directory):
    # -I -S excludes the author's installed packages and the recipient's user site.
    support = str(Path(__file__).resolve().parent)
    sys.path.insert(0, support)
    from _xll_functions import _command_label, _register
    import site
    site.addsitedir(str(directory / "packages"))
    sys.path.insert(1, str(directory / "src"))
    project = json.loads((directory / "project.json").read_text(encoding="utf-8"))
    legacy = set(project.get("legacyCommands", ()))
    commands, entries, functions = {}, [], {}
    for relative in project["entries"]:
        parts = list(Path(relative).with_suffix("").parts)
        if parts[-1] == "__init__":
            parts.pop()
        module = importlib.import_module(".".join(parts))
        expected = (directory / "src" / relative).resolve()
        if Path(module.__file__).resolve() != expected:
            raise ValueError(f"Entry module {relative} conflicts with another Python module")
        local = [fn for fn in vars(module).values()
                 if inspect.isfunction(fn) and fn.__module__ == module.__name__]
        for fn in dict.fromkeys([*local, *getattr(module, "__excel_exports__", ())]):
            _register(fn, functions, entries)
        for name, fn in vars(module).items():
            if not inspect.isfunction(fn) or fn.__module__ != module.__name__:
                continue
            target = module.__name__ + "." + name
            marked = name == fn.__name__ and hasattr(fn, "__excel_macro__")
            if not marked and (name.startswith("_") or target not in legacy):
                continue
            try:
                _command_label(fn)
            except (TypeError, ValueError):
                continue
            commands[target] = fn
    return functions, commands, entries


def serve(root):
    # Keep stdout from Python, native extensions and subprocesses off the response pipe.
    protocol = os.fdopen(os.dup(sys.stdout.fileno()), "w", encoding="utf-8", newline="\n")
    os.dup2(sys.stderr.fileno(), sys.stdout.fileno())
    incoming = sys.stdin.buffer
    # Only this loop consumes requests, including when a user script calls input().
    sys.stdin = open(os.devnull, encoding="utf-8")
    # Keep relative writes private and separate from the importable code. Two Excel
    # instances may open the same add-in for the first time concurrently.
    work = root / "work"
    if not work.exists():
        with tempfile.TemporaryDirectory(dir=root) as temp:
            seed = Path(temp) / "work"
            shutil.copytree(root / "src", seed)
            try:
                seed.rename(work)
            except FileExistsError:
                pass
    os.chdir(work)
    functions, commands, _ = load(root)
    from _xll_functions import _call, _encode, _invoke, _run_command
    from _xll_streams import Streams
    subscriptions = Streams()
    from _xll_application import Session, active
    from concurrent.futures import ThreadPoolExecutor
    write = threading.Lock()

    def respond(packet):
        request = packet["request"]
        try:
            op = request["op"]
            value = None
            if op == "command":
                session = Session(request["excelPipe"])
                try:
                    with active(session):
                        _run_command(commands[request["name"]])
                    value = None
                finally:
                    session.close()
            elif op == "call":
                value = _invoke(functions[request["name"].upper()], request["args"])
            elif op == "streamOpen":
                generator = _call(functions[request["name"].upper()], request["args"])
                subscriptions.open(request["id"], generator, _encode)
            elif op == "streamRead":
                value = subscriptions.read()
            elif op == "streamClose":
                subscriptions.close(request["id"])
            else:
                raise ValueError(f"Unknown worker operation: {op}")
            reply = json.dumps({"id": packet["id"], "value": value}, allow_nan=False)
        except BaseException:
            reply = json.dumps({"id": packet["id"], "failure": traceback.format_exc()})
        with write:
            protocol.write(reply + "\n")
            protocol.flush()

    # Stream reads and ribbon commands may wait for Excel while it calculates a formula.
    with ThreadPoolExecutor() as formulas, ThreadPoolExecutor() as actions, ThreadPoolExecutor(1) as streams:
        for line in incoming:
            packet = json.loads(line)
            op = packet["request"]["op"]
            pool = streams if op == "streamRead" else actions if op == "command" else formulas
            pool.submit(respond, packet)
        os._exit(0)  # Excel closed the transport; user threads may be waiting indefinitely.


if __name__ == "__main__":
    root = Path(sys.argv[2])
    if sys.argv[1] == "inspect":
        _, commands, entries = load(root)
        Path(sys.argv[3]).write_text(json.dumps(dict(commands=sorted(commands), functions=entries)), encoding="utf-8")
    else:
        serve(root)
