"""Exercise console behavior through the same JSON entry point as the XLL."""
import io
import json
from pathlib import Path
import sys
import tempfile
import threading
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "python"))
from _xll_host import dispatch


class ConsoleTest(unittest.TestCase):
    def host(self, op, **args):
        reply = json.loads(dispatch(json.dumps(dict(op=op, **args))))
        self.assertNotIn("failure", reply)
        return reply["value"]

    def call(self, action, **args):
        return self.host("repl", session="test", action=action, args=args)

    def setUp(self):
        self.call("open")
        self.addCleanup(self.call, "close")

    def test_state_continuation_and_error_recovery(self):
        self.assertTrue(self.call("run", source="def add(n):")["incomplete"])
        self.call("run", source="def add(n):\n    return n + 2\nadd(3)")
        self.assertEqual(self.call("drain")[0]["value"]["text"], "5")
        self.call("run", source="print(_)\n1/0")
        events = self.call("drain")
        self.assertEqual(events[0]["value"], "5")
        self.assertIn("ZeroDivisionError", events[-1]["value"])
        self.call("run", source="add(4)")
        self.assertEqual(self.call("drain")[0]["value"]["text"], "6")

    def test_worksheet_exports_and_session_lifetime(self):
        def invoke(*args):
            return self.host("call", name="XLLPY.CONSOLE.SCALE", args=args)

        self.call("run", source="""from statistics import mean
def _helper(n): return n + 1
def unsupported(*items): return items
rate = 2
def scale(n: int, factor=3): return _helper(n) * factor * rate""")
        self.assertIn("fixed positional signature", self.call("drain")[0]["value"])
        exports = self.host("worksheetFunctions")
        self.assertEqual([f["name"] for f in exports], ["XLLPY.console.scale"])
        self.assertEqual(invoke(3.5, {"missing": True}), 24)
        self.assertNotIn("worksheetFormulas", self.call("run", source="rate = 1"))
        self.assertEqual(invoke(3.5), 12)
        self.call("clear")
        self.assertEqual(self.call("run", source="def scale(n): return [[n, rate]]")["worksheetFormulas"],
                         ["=XLLPY.console.scale(n)"])
        self.assertEqual(invoke(7), [[7, 1]])
        self.assertTrue(self.call("detach"))
        self.assertEqual(invoke(7), {"error": 29})
        # An old window can finish executing after its replacement has opened.
        self.host("repl", session="next", action="open")
        self.addCleanup(self.host, "repl", session="next", action="close")
        self.call("run", source="def old(): return 99")
        self.call("close")
        self.assertEqual(self.host("worksheetFunctions"), [])

    def test_startup_refresh_and_console_name_collision(self):
        self.call("run", source="def scale(n): return 2*n\ndef answer(): return 42")
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder, "functions.py")
            path.write_text("from xll_python import function\n@function(name='XLLPY.CONSOLE.SCALE')\n"
                            "def triple(n): return 3*n\n", encoding="utf-8")
            self.addCleanup(self.host, "discover", folders=[], packages=[], refresh=True)
            discovered = self.host("discover", folders=[folder], packages=[], refresh=True)
            self.assertEqual(discovered["errors"], [])
            self.assertEqual({f["name"] for f in discovered["functions"]},
                             {"XLLPY.CONSOLE.SCALE", "XLLPY.console.answer"})
            self.assertEqual(self.host("call", name="xllpy.console.scale", args=[4]), 12)
            self.host("discover", folders=[], packages=[], refresh=True)
            self.assertEqual(self.host("call", name="xllpy.console.scale", args=[4]), 8)

    def test_live_objects_paging_and_cycles(self):
        self.call("run", source="a = list(range(60))\na[0] = a\nprint('items', a, end='!')")
        events = self.call("drain")
        self.assertEqual([e["kind"] for e in events], ["stdout", "stdout", "object", "stdout"])
        handle = events[2]["value"]["handle"]
        self.call("run", source="a[1] = 80")
        first = self.call("children", handle=handle, offset=0)
        self.assertTrue(first["more"])
        self.assertEqual(first["items"][1]["text"], "80")
        cycle = first["items"][0]["handle"]
        self.assertEqual(self.call("children", handle=cycle, offset=50)["items"][-1]["text"], "59")
        self.call("clear")
        self.call("run", source="a[1]")
        self.assertEqual(self.call("drain")[0]["value"]["text"], "80")

    def test_static_completion_and_field_inspection(self):
        self.call("run", source="""from __future__ import annotations
from dataclasses import dataclass
@dataclass(slots=True)
class Record:
    data: list
    @property
    def danger(self):
        raise AssertionError("getter ran")
r = Record([42])
r""")
        handle = self.call("drain")[0]["value"]["handle"]
        self.assertEqual(self.call("complete", token="r.data.ap"), ["append"])
        self.assertEqual(self.call("complete", token="r.dan"), ["danger"])
        fields = self.call("children", handle=handle, offset=0)["items"]
        self.assertIn("data", [item["label"] for item in fields])

    def test_thread_local_output_and_explicit_destinations(self):
        self.call("run", source="""from contextlib import redirect_stdout
from io import StringIO
import sys
stream = StringIO()
with redirect_stdout(stream):
    print([123])
print(42, file=stream)
print(stream.getvalue(), end="")""")
        self.assertEqual(self.call("drain"), [dict(kind="stdout", value="[123]\n42\n")])
        # The script runner and worksheet threads share the interpreter's stream proxies.
        entered, release = threading.Event(), threading.Event()
        import _xll_host
        console = _xll_host._consoles["test"]
        console.namespace.update(entered=entered, release=release)
        result = []
        command = threading.Thread(target=lambda: result.append(self.call("run", source=
            "print('before'); entered.set(); release.wait(3); print('after', file=sys.stderr)")))
        command.start()
        try:
            self.assertTrue(entered.wait(2))
            self.assertEqual(self.call("drain")[0]["value"], "before")
            outside = io.StringIO()
            _xll_host._capture.stream = outside
            try:
                print("outside")
            finally:
                del _xll_host._capture.stream
            self.assertEqual(outside.getvalue(), "outside\n")
        finally:
            release.set()
            command.join(3)
        self.assertEqual(result, [dict(incomplete=False)])
        self.assertEqual(self.call("drain"), [dict(kind="stderr", value="after"), dict(kind="stderr", value="\n")])


if __name__ == "__main__":
    unittest.main()
