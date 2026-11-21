"""VBA-style access to the Excel instance serving this command."""
from contextlib import contextmanager
import threading
import io
import json
import weakref


_current = threading.local()


class Session:
    def __init__(self, pipe):
        self.pipe = pipe
        self.stream = self.reader = None
        self.released = []

    def request(self, target, action, member, args):
        if self.stream is None:
            self.stream = open(r"\\.\pipe" + "\\" + self.pipe, "r+b", buffering=0)
            self.reader = io.BufferedReader(self.stream)
        def encode(arg):
            if isinstance(arg, _Object):
                if arg._session is not self:
                    raise ValueError("Excel object belongs to another command session")
                return {"handle": arg._handle}
            raise TypeError(f"Unsupported Excel argument: {type(arg).__name__}")
        packet = json.dumps(dict(target=target, action=action, member=member, args=args,
                                 release=self.released), default=encode, allow_nan=False)
        self.stream.write((packet + "\n").encode())
        self.released.clear()
        reply = self.reader.readline()
        if not reply:
            raise RuntimeError("The Excel command session closed")
        result = json.loads(reply)
        if "error" in result:
            raise RuntimeError(result["error"])
        value = result["value"]
        if isinstance(value, dict) and "handle" in value:
            return _Object(self, value["handle"])
        return value

    def close(self):
        if self.reader:
            self.reader.close()


@contextmanager
def active(session):
    previous = getattr(_current, "session", None)
    _current.session = session
    try:
        yield
    finally:
        _current.session = previous


class _Object:
    def __init__(self, session, handle):
        object.__setattr__(self, "_session", session)
        object.__setattr__(self, "_handle", handle)
        if handle:
            weakref.finalize(self, session.released.append, handle)

    def _request(self, action, name, args):
        session = getattr(_current, "session", None)
        if session is None or (self._handle and self._session is not session):
            raise RuntimeError("Application requires an active Excel script, console or ribbon command")
        return session.request(self._handle, action, name, args)

    def __getattr__(self, name):
        if name.startswith("_"):
            raise AttributeError(name)
        return self._request("get", name, ())

    def __setattr__(self, name, value):
        self._request("set", name, (value,))

    def __call__(self, *indices):
        return self._request("get", "Item", indices)

    def Range(self, *address):
        return self._request("get", "Range", address)

    def Call(self, method, *args):
        """Invoke an Excel method, e.g. sheet.Call("Activate")."""
        return self._request("call", method, args)


Application = _Object(None, 0)
