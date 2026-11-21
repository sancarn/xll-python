"""Build two offline projects, then exercise their isolated packages, formulas and ribbon in Excel."""
import base64
import ctypes
import json
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
import time
import uuid
import zipfile

from excel_smoke import ROOT, connect, wait_for
import psutil
import win32gui
import win32ui
from PIL import Image
from pywinauto import Desktop


def wheel(folder, name, version, source, requires=""):
    path = folder / f"{name}-{version}-py3-none-any.whl"
    info = f"{name}-{version}.dist-info/"
    files = {name+".py": source,
             info+"METADATA": f"Metadata-Version: 2.1\nName: {name}\nVersion: {version}\n{requires}\n",
             info+"WHEEL": "Wheel-Version: 1.0\nGenerator: fixture\nRoot-Is-Purelib: true\nTag: py3-none-any\n"}
    files[info+"RECORD"] = "".join(f"{k},,\n" for k in [*files, info+"RECORD"])
    with zipfile.ZipFile(path, "w") as archive:
        for key, value in files.items():
            archive.writestr(key, value)
    return path


def project(folder, version):
    label = "First" if version == 1 else "Second"
    folder.mkdir()
    wheel(folder, "xll_fixture_dep", version, f"VALUE={version}\n")
    primary = wheel(folder, "xll_fixture", version, "from xll_fixture_dep import VALUE\n",
                    f"Requires-Dist: xll_fixture_dep=={version}\n")
    (folder / "helper.py").write_text(f"OFFSET={10*version}\n")
    (folder / "reports.py").write_text(f"""import os
from xll_python import Application, function
from helper import OFFSET
from xll_fixture import VALUE

@function(name="Pkg.{label}")
def price(x):
    return VALUE + OFFSET + x

@function(name="Pkg.{label}Range")
def numbers(rows):
    return [[VALUE + v for v in row] for row in rows]

@function(name="Pkg.{label}Pid")
def worker_pid():
    return os.getpid()

@function(name="Pkg.{label}Async")
async def async_value():
    return VALUE

@function(name="Pkg.{label}Stream")
def stream_value():
    yield VALUE

def refresh():
    cell = Application.ActiveSheet.Range("B{version}")
    cell.Value2 = "{label} " + str(VALUE)
""")
    png = "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAN0lEQVR4nO3QQREAMAgDQYoOhKOrhspUBZ+NgcvsudUvFpebcQcIECBAgAABAgQIECBAgACBLzCwaALFNKkexgAAAABJRU5ErkJggg=="
    (folder / "icon.png").write_bytes(base64.b64decode(png))
    tabs = [dict(label=label, groups=[dict(label="Reports", buttons=[dict(label="Refresh " + label,
            command="reports.refresh", image="icon.png", large=True)])])]
    manifest = folder / "project.xllproject.json"
    manifest.write_text(json.dumps(dict(name=label, id=str(uuid.uuid4()), entries=["reports.py"],
                             files=["helper.py"], packages=[primary.name], tabs=tabs)))
    return manifest


def button_icon(hwnd, control):
    # Render the actual Excel window, even on a secondary or obscured desktop.
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    dc = win32gui.GetWindowDC(hwnd)
    source = win32ui.CreateDCFromHandle(dc)
    target = source.CreateCompatibleDC()
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(source, right-left, bottom-top)
    previous = target.SelectObject(bmp)
    try:
        assert ctypes.windll.user32.PrintWindow(hwnd, target.GetSafeHdc(), 2)
        image = Image.frombytes("RGB", (right-left, bottom-top), bmp.GetBitmapBits(True), "raw", "BGRX")
        rect = control.rectangle()
        return image.crop((rect.left-left, rect.top-top, rect.right-left, rect.bottom-top))
    finally:
        target.SelectObject(previous)
        win32gui.DeleteObject(bmp.GetHandle())
        target.DeleteDC()
        win32gui.ReleaseDC(hwnd, dc)


def main():
    template = ROOT / "dist/PythonExcel.xll"
    with tempfile.TemporaryDirectory(dir=ROOT/".build", prefix="compiled-test-") as tmp:
        base = Path(tmp)
        build_env = dict(os.environ, PIP_NO_INDEX="1")
        home = base/"build-runtime"
        with zipfile.ZipFile(ROOT/".build/runtime.zip") as archive:
            archive.extractall(home)
        python = home / "python.exe"
        subprocess.run([python, "-I", "-m", "ensurepip"], check=True, env=build_env, stdout=subprocess.DEVNULL)
        outputs = []
        for version, label in enumerate(("First", "Second"), 1):
            source = project(base/label, version)
            output = base/(label+".xll")
            if version == 1:
                author_env = dict(build_env, PIP_FIND_LINKS=str(source.parent))
                subprocess.run([python, "-I", "-m", "pip", "install", "xll_fixture==1"], check=True, env=author_env)
                subprocess.run([python, "-I", ROOT/"tests/compiler_designer.py", source, template, output],
                               check=True, env=author_env)
            else:
                wheelhouse = base / "external wheelhouse"
                wheelhouse.mkdir()
                for path in source.parent.glob("*.whl"):
                    path.rename(wheelhouse/path.name)
                wheel = next(wheelhouse.glob("xll_fixture-*.whl"))
                cached = base / wheel.name
                wheel.rename(cached)
                try:
                    wheel.symlink_to(cached)
                except OSError:  # Developer Mode is optional on the Windows test machine.
                    shutil.copyfile(cached, wheel)
                subprocess.run([python, "-I", ROOT/"tests/compiler_designer.py", source, template, output, wheel],
                               check=True, env=build_env)
                subprocess.run([python, "-I", ROOT/"python/_xll_compile.py", source, template, output],
                               check=True, env=build_env)
            shutil.rmtree(source.parent)
            outputs.append(output)
        # Same source and dependencies, distinct saved projects and writable state.
        state_dir = base / "state-project"
        state_dir.mkdir()
        (state_dir / "state.py").write_text(
            "from pathlib import Path\nfrom xll_python import Application\n"
            "def increment():\n    counter = Path('count.txt')\n"
            "    count = int(counter.read_text()) + 1 if counter.exists() else 1\n"
            "    counter.write_text(str(count))\n    Path('state.py').write_text('working copy')\n"
            "    assert 'def increment' in Path(__file__).read_text()\n"
            "    Application.ActiveSheet.Range('D1').Value2 = count\n")
        for name in ["StateA", "StateB"]:
            manifest = state_dir / "project.json"
            manifest.write_text(json.dumps(dict(id=str(uuid.uuid4()), name=name, entries=["state.py"],
                files=[], packages=[], tabs=[dict(label=name, groups=[dict(label="State", buttons=[
                    dict(label="Increment", command="state.increment", imageMso="MacroPlay")])])])))
            output = base / (name+".xll")
            subprocess.run([python, "-I", ROOT/"python/_xll_compile.py", manifest, template, output],
                           env=build_env, check=True)
            outputs.append(output)
        shutil.rmtree(home)
        book = base/"test.csv"
        book.write_text("data\n")
        profile = base/"recipient"
        excel = Path(os.environ["ProgramFiles"])/"Microsoft Office/root/Office16/EXCEL.EXE"
        proc = subprocess.Popen([excel, "/x", "/automation", book],
                                env=dict(build_env, XLL_PYTHON_HOME=str(profile)))
        app = None
        try:
            app = wait_for(lambda: connect(proc.pid, book))
            app.DisplayAlerts = False
            app.Visible = True
            assert app.RegisterXLL(str(template))
            for output in outputs:
                assert app.RegisterXLL(str(output))
            assert not (profile/"published").exists(), "Excel started Python before a formula or click"
            sheet = app.ActiveSheet
            for i, (label, expected) in enumerate([("First", 18), ("Second", 29)], 1):
                cell = sheet.Range(f"A{i}")
                cell.Formula2 = f"=Pkg.{label}(7)"
                assert cell.Value2 == expected, (label, cell.Text)
                sheet.Range(f"C{i}").Formula2 = f"=Pkg.{label}Pid()"
                for col, suffix in [("J","Async"), ("K","Stream")]:
                    topic = sheet.Range(f"{col}{i}")
                    topic.Formula2 = f"=Pkg.{label}{suffix}()"
                    wait_for(lambda: topic.Value2 == i)
            pids = [int(sheet.Range(f"C{i}").Value2) for i in (1,2)]
            assert len(set([proc.pid, *pids])) == 3
            main = Desktop(backend="uia").window(handle=app.Hwnd)
            for i, label in enumerate(("First", "Second"), 1):
                main.child_window(title=label, control_type="TabItem").select()
                control = main.child_window(title="Refresh "+label, control_type="Button").wrapper_object()
                def icon_visible():
                    image = button_icon(app.Hwnd, control)
                    return sum(count for count, (r,g,b) in image.getcolors(image.width*image.height)
                               if r > 175 and g < 65 and 120 < b < 180) > 50
                wait_for(icon_visible, 10)
                control.invoke()
                wait_for(lambda: sheet.Range(f"B{i}").Value2 == f"{label} {i}")
            for label, expected in [("StateA", 1), ("StateB", 1), ("StateA", 2)]:
                sheet.Range("D1").ClearContents()
                main.child_window(title=label, control_type="TabItem").select()
                main.child_window(title="Increment", control_type="Button").invoke()
                wait_for(lambda: sheet.Range("D1").Value2 == expected)
            app.Calculation = -4135
            sheet.Range("E1:E10000").Value2 = tuple((1,) for _ in range(10000))
            sheet.Range("F1:F10000").FormulaR1C1 = "=Pkg.First(RC[-1])"
            t = time.perf_counter()
            sheet.Range("F1:F10000").Calculate()
            scalar = time.perf_counter()-t
            assert sheet.Range("F10000").Value2 == 12
            sheet.Range("G1").Formula2 = "=Pkg.FirstRange(E1:E10000)"
            t = time.perf_counter()
            sheet.Range("G1").Calculate()
            print(f"10,000 scalar formulas {scalar:.3f}s; one range {time.perf_counter()-t:.3f}s", flush=True)
            assert sheet.Range("G10000").Value2 == 2
            print("PASS: workbench and two standalone XLLs, lazy isolated workers, conflicting transitive wheels, custom ribbon buttons and workbook edits", flush=True)
        finally:
            if app is not None:
                app.DisplayAlerts = False
                app.Quit()
            try:
                proc.wait(15)
            except subprocess.TimeoutExpired:
                proc.kill()
                proc.wait()
        assert all(not psutil.pid_exists(pid) for pid in pids)


if __name__ == "__main__":
    main()
