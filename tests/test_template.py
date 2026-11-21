"""Exercise first download, offline reuse and integrity failures without contacting a release."""
from hashlib import sha256
from io import BytesIO
import json
import os
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "python"))
import _xll_compile as compiler


class TemplateTests(unittest.TestCase):
    def test_pinned_download_and_offline_cache(self):
        data = b"runtime template"
        pin = dict(tag="v0.2.17", sha256=sha256(data).hexdigest())
        with tempfile.TemporaryDirectory() as tmp, patch.dict(os.environ, XLL_PYTHON_HOME=tmp), \
                patch.object(compiler, "resource_bytes", return_value=json.dumps(pin)), \
                patch.object(compiler.urllib.request, "urlopen") as download:
            workbench = Path(tmp) / "PythonExcel.xll"
            output = Path(tmp) / "Example.xll"
            download.return_value = BytesIO(data)
            self.assertEqual(compiler.optimised_template(workbench, output), data)
            download.assert_called_once_with(
                "https://github.com/sancarn/xll-python/releases/download/v0.2.17/PythonExcelRuntime.xll", timeout=90)
            download.reset_mock()
            self.assertEqual(compiler.optimised_template(workbench, output), data)
            download.assert_not_called()
            cache, = Path(tmp).rglob("*.xll")
            cache.write_bytes(b"corrupt")
            download.return_value = BytesIO(b"different build")
            with self.assertRaisesRegex(ValueError, "SHA-256 verification"):
                compiler.optimised_template(workbench, output)
            self.assertEqual(cache.read_bytes(), b"corrupt")
            def concurrent_replace(candidate, target):
                target.write_bytes(candidate.read_bytes())
                raise PermissionError("another build is reading the published file")

            download.return_value = BytesIO(data)
            with patch.object(compiler.os, "replace", side_effect=concurrent_replace):
                self.assertEqual(compiler.optimised_template(workbench, output), data)
            self.assertEqual(cache.read_bytes(), data)


if __name__ == "__main__":
    unittest.main()
