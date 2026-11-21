"""Check the CI-built PE resource and packaged support files without desktop Excel."""
import ctypes
from ctypes import wintypes
import hashlib
import io
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import uuid
import zipfile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "python"))
from _xll_compile import resource_bytes


def main():
    root = ROOT
    workbench = root / "dist/PythonExcel.xll"
    runtime_template = root / "dist/PythonExcelRuntime.xll"
    rcdata = ctypes.cast(10, wintypes.LPCWSTR)
    pack = root / ".build/pack"
    for name, source in [("PYTHON_RUNTIME", "runtime.zip"), ("EXCEL_IDE", "ide.zip")]:
        assert resource_bytes(workbench, name, rcdata) == (root / ".build" / source).read_bytes()
    assembly = resource_bytes(workbench, "PYTHONEXCEL", "ASSEMBLY")
    assert assembly == (pack / "PythonExcel.dll").read_bytes()
    for name in ("updater.exe", "support.zip"):
        assert (pack / name).read_bytes() in assembly
    with zipfile.ZipFile(pack / "support.zip") as archive:
        for path in ("LICENSE", "LICENSE-PACKAGE-EXCEPTIONS.txt", "examples/LICENSE",
                     "notices/Excel-DNA.txt", "notices/Theia-LICENSE-EPL", "notices/Theia-NOTICE.md",
                     "notices/Third-party-sources.md"):
            assert archive.read(path) == (root / path).read_bytes(), path
    pin = json.loads(resource_bytes(workbench, "RUNTIME_TEMPLATE", rcdata))
    assert pin["sha256"] == hashlib.sha256(runtime_template.read_bytes()).hexdigest()
    assert pin["tag"] == (pack / "release.txt").read_text()
    compiled = resource_bytes(runtime_template, "PYTHONEXCELRUNTIME", "ASSEMBLY")
    assert compiled == (pack / "PythonExcelRuntime.dll").read_bytes()
    assert (pack / "published-support.zip").read_bytes() in compiled
    assert b"WebView2" not in compiled and b"updater.exe" not in compiled
    assert b"WebView2" not in resource_bytes(runtime_template, "__MAIN__", "DNA")
    assert resource_bytes(runtime_template, "PYTHON_RUNTIME", rcdata) == (root / ".build/published-runtime.zip").read_bytes()
    for name, kind in [("EXCEL_IDE", rcdata), ("PYTHONEXCEL", "ASSEMBLY"),
                       ("MICROSOFT.WEB.WEBVIEW2.CORE", "ASSEMBLY"), ("MICROSOFT.WEB.WEBVIEW2.WINFORMS", "ASSEMBLY")]:
        try:
            resource_bytes(runtime_template, name, kind)
        except ValueError:
            pass
        else:
            raise AssertionError(name)
    with tempfile.TemporaryDirectory() as tmp:
        folder = Path(tmp)
        with zipfile.ZipFile(root / ".build/published-runtime.zip") as archive:
            assert "LICENSE.txt" in archive.namelist()
            assert not any(s.startswith(("include/", "libs/", "Lib/test/")) for s in archive.namelist())
            archive.extractall(folder / "python")
        with zipfile.ZipFile(pack / "published-support.zip") as archive:
            assert set(archive.namelist()) == {"_xll_worker.py", "_xll_functions.py", "_xll_application.py",
                "_xll_streams.py", "xll_python.py", "LICENSE", "LICENSE-PACKAGE-EXCEPTIONS.txt",
                "notices/Excel-DNA.txt", "notices/Third-party-sources.md"}
            archive.extractall(folder / "support")
        python = folder / "python/python.exe"
        (folder / "example.py").write_text("from xll_python import function\n@function\ndef twice(x): return 2 * x\n")
        project = folder / "project.json"
        settings = dict(name="Example", id=str(uuid.uuid4()), entries=["example.py"], files=[], packages=[], tabs=[])
        project.write_text(json.dumps(settings))
        standard = folder / "Standard.xll"
        compiler = [python, "-I", root / "python/_xll_compile.py", project, workbench]
        subprocess.run([*compiler, standard], check=True)
        assert resource_bytes(standard, "PYTHONEXCEL", "ASSEMBLY") == assembly
        project.write_text(json.dumps(dict(settings, optimise=True)))
        output = folder / "Example.xll"
        subprocess.run([*compiler, output], check=True)
        assert output.stat().st_size < standard.stat().st_size
        print(f"Same project: {standard.stat().st_size:,} bytes standard, {output.stat().st_size:,} optimised")
        assert resource_bytes(output, "PYTHONEXCELRUNTIME", "ASSEMBLY") == compiled
        manifest = json.loads(resource_bytes(output, "PYTHON_MANIFEST", rcdata))
        assert manifest["functions"][0]["name"] == "twice" and manifest["ribbon"] == ""
        with zipfile.ZipFile(io.BytesIO(resource_bytes(output, "PYTHON_PROJECT", rcdata))) as archive:
            archive.extractall(folder / "project")
        with subprocess.Popen([python, "-I", "-S", folder / "support/_xll_worker.py", "serve", folder / "project"],
                              stdin=subprocess.PIPE, stdout=subprocess.PIPE, text=True) as proc:
            proc.stdin.write(json.dumps(dict(id=1, request=dict(op="call", name="twice", args=[21]))) + "\n")
            proc.stdin.flush()
            assert json.loads(proc.stdout.readline()) == {"id": 1, "value": 42}
            proc.stdin.close()
            assert proc.wait(timeout=15) == 0
    with zipfile.ZipFile(root/".build/runtime.zip") as archive:
        assert archive.testzip() is None
        assert b"PYTHON SOFTWARE FOUNDATION LICENSE" in archive.read("LICENSE.txt")
    print("Verified workbench and runtime-only exports", hashlib.sha256((root/"dist/PythonExcel.xll").read_bytes()).hexdigest())


if __name__ == "__main__":
    main()
