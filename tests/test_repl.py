"""Exercise console behavior through the same JSON entry point as the XLL."""
import io
import json
from pathlib import Path
import sys
import threading
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "python"))
from _xll_host import dispatch


class ConsoleTest(unittest.TestCase):
    def call(self, action, **args):
        reply = json.loads(dispatch(json.dumps(dict(op="repl", session="test", action=action, args=args))))
        self.assertNotIn("failure", reply)
        return reply["value"]

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
