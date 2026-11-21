"""Exercise macro discovery, refresh and execution through the XLL host protocol."""
import json
from pathlib import Path
import sys
import tempfile
import unittest
import uuid

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "python"))
from _xll_compile import command_names, publish
from _xll_host import dispatch


def request(op, **data):
    return json.loads(dispatch(json.dumps(dict(op=op, **data))))


class MacroTest(unittest.TestCase):
    @unittest.skipUnless(sys.version_info >= (3, 12), "Python 3.12 introduced generic functions")
    def test_generic_generator_is_excluded(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            (root / "generic.py").write_text(
                "from xll_python import macro\n"
                "@macro\ndef rows[T](): yield 1\n"
                "@macro\nasync def events[T](): yield 1\n"
                "@macro\ndef report[T](): pass\n")
            self.assertEqual(command_names(root, ["generic.py"]), [dict(label="report", target="generic.report")])

    def test_invalid_macro_leaves_other_startup_exports_available(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            entry = root / "mixed.py"
            entry.write_text(
                "from xll_python import function, macro\n"
                "@macro\ndef missing(arg): pass\n"
                "@macro\ndef stream(): yield from range(2)\n"
                "@macro\nasync def async_stream(): yield 1\n"
                "@macro\ndef nested_default():\n    def inner(value=(yield 1)): pass\n"
                "@macro(name=42)\ndef invalid_label(): pass\n"
                "@macro\ndef report():\n    def rows(): yield 1\n    print(sum(rows()))\n"
                "@function\ndef total(): return 7\n")
            self.addCleanup(request, "discover", folders=[], packages=[], refresh=True)
            catalog = request("discover", folders=[folder], packages=[])["value"]
            self.assertEqual([f["name"] for f in catalog["functions"]], ["total"])
            self.assertEqual([m["label"] for m in catalog["macros"]], ["report"])
            self.assertEqual(len(catalog["errors"]), 5)
            self.assertEqual(command_names(root, [entry.name]), [dict(label="report", target="mixed.report")])

    def test_new_binding_requires_decorator_at_compile_time(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            (root / "actions.py").write_text("def extract(): pass\n")
            button = dict(label="Extract CSV", command="actions.extract", macro=True)
            manifest = root / "project.json"
            manifest.write_text(json.dumps(dict(id=str(uuid.uuid4()), name="Test", entries=["actions.py"],
                files=[], packages=[], tabs=[dict(label="Reports", groups=[dict(label="CSV", buttons=[button])])])))
            with self.assertRaisesRegex(ValueError, "actions.extract must be marked with @macro"):
                publish(manifest, root / "template.xll", root / "output.xll")

    def test_cached_macros_survive_failed_refresh(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            (root / "_library.py").write_text("from xll_python import macro\n@macro\ndef imported(): pass\n")
            entry = root / "functions.py"
            entry.write_text(
                "from _library import imported\nimport xll_python as xl\n"
                "from xll_python import macro as action\n"
                "@xl.macro(name='Extract CSV')\nasync def extract(limit=10): print(limit)\n"
                "@action\ndef update(): print('updated')\n"
                "def helper(): pass\n")
            self.addCleanup(request, "discover", folders=[], packages=[], refresh=True)
            def refresh():
                return request("discover", folders=[folder], packages=[], refresh=True)
            catalog = refresh()["value"]
            self.assertEqual(catalog["errors"], [])
            self.assertEqual(catalog["functions"], [])
            self.assertEqual([m["label"] for m in catalog["macros"]], ["Extract CSV", "update"])
            self.assertEqual(command_names(root, [entry.name]), [
                dict(label="Extract CSV", target="functions.extract"), dict(label="update", target="functions.update")])
            target = catalog["macros"][0]["target"]
            entry.write_text("from xll_python import macro\n@macro\ndef broken(required): pass\n")
            self.assertIn("callable without arguments", refresh()["failure"])
            self.assertEqual(request("macro", target=target), {"value": "10\n"})
            entry.write_text("def helper(): pass\n")
            self.assertEqual(refresh()["value"]["macros"], [])
            self.assertIn("failure", request("macro", target=target))


if __name__ == "__main__":
    unittest.main()
