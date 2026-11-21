"""End-to-end smoke test for 64-bit desktop Excel. Requires pywin32 (test driver only)."""
import json
import os
import shutil
from pathlib import Path
import subprocess
import sys
import tempfile
import time
sys.coinit_flags = 2

import pythoncom
import win32com.client
import win32con
import win32gui
import win32process

ROOT = Path(__file__).resolve().parents[1]


def wait_for(probe, timeout=45):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        value = probe()
        if value:
            return value
        time.sleep(.15)
    raise TimeoutError(probe.__name__)


def connect(pid, workbook):
    table = pythoncom.GetRunningObjectTable()
    ctx = pythoncom.CreateBindCtx(0)
    for moniker in table:
        if moniker.GetDisplayName(ctx, None).casefold() == str(workbook).casefold():
            book = win32com.client.dynamic.Dispatch(table.GetObject(moniker).QueryInterface(pythoncom.IID_IDispatch))
            app = book.Application
            if win32process.GetWindowThreadProcessId(app.Hwnd)[1] == pid:
                return app


def verify_streams(app, sheet, scripts):
    def number(cell):
        value = cell.Value2
        return value if isinstance(value, float) else None

    live = sheet.Range("N1")
    live.Formula2 = "=XLLPY.live_random()"
    first = wait_for(lambda: number(live))
    sheet.Range("O1").Formula2 = "=N1*2"
    def dependent_updated():
        value, doubled = sheet.Range("N1:O1").Value2[0]
        return isinstance(value, float) and value != first and doubled == 2 * value
    wait_for(dependent_updated)

    a, b, c = (sheet.Range(address) for address in ("N2", "N3", "N4"))
    a.Formula2 = b.Formula2 = '=PY.SHARED("a")'
    c.Formula2 = '=PY.SHARED("b")'
    instance = wait_for(lambda: number(a))
    wait_for(lambda: number(b) == instance)
    assert wait_for(lambda: number(c)) != instance
    same = sheet.Range("N5")
    same.Formula2 = "=PY.PERCELL()-PY.PERCELL()"
    wait_for(lambda: number(same) == 0)
    cell = sheet.Range("N6")
    cell.Formula2 = "=PY.PERCELL()"
    cell_id = wait_for(lambda: number(cell))
    other = app.Workbooks.Add()
    try:
        twin = other.Worksheets(1).Range("N6")
        twin.Formula2 = "=PY.PERCELL()"
        assert wait_for(lambda: number(twin)) != cell_id
        twin.Formula2 = '=PY.SHARED("a")'
        wait_for(lambda: number(twin) == instance)
    finally:
        other.Close(False)
    app.CalculateFull()
    wait_for(lambda: number(cell) == cell_id and number(a) == instance)
    a.ClearContents()
    closed = scripts / f"closed-{int(instance)}"
    assert b.Value2 == instance and not closed.exists()
    b.ClearContents()
    wait_for(closed.exists)

    for address, kind, expected in (("N7", "empty", 0), ("N8", "error", "#VALUE!"), ("N9", "na", "#N/A")):
        target = sheet.Range(address)
        target.Formula2 = f'=PY.FINITE("{kind}")'
        wait_for(lambda: target.Text == expected if isinstance(expected, str) else target.Value2 == expected)
    sheet.Range("P1").Formula2 = '=PY.FINITE("array")'
    wait_for(lambda: sheet.Range("P1:Q2").Value2 == ((1., 2.), (3., 4.)))
    sheet.Range("P1").ClearContents()
    sheet.Range("N1:O9").ClearContents()
    print("PASS: live updates, shared and per-cell generators, recalc, disconnect, completion and errors.", flush=True)


def main():
    if "--ui" in sys.argv:
        from excel_ui import verify
    xll = (ROOT / "dist/PythonExcel.xll").resolve(strict=True)
    excel = Path(os.environ["ProgramFiles"]) / "Microsoft Office/root/Office16/EXCEL.EXE"
    # Keep the editor fixture outside AppData, which Chromium classifies as a
    # protected system directory. Release it after both Excel and WebView2 exit.
    parent = ROOT / ".build" if os.environ.get("XLL_WEB_TEST") else None
    if parent:
        parent.mkdir(exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="xll-python-test-",dir=parent) as tmp:
        base = Path(tmp)
        standalone = base / "distribution" / xll.name
        standalone.parent.mkdir()
        shutil.copyfile(xll, standalone)
        scripts, packages, profile = (base / n for n in ("scripts","packages","profile"))
        for d in (scripts,packages,profile):
            d.mkdir()
        (base/"book.csv").write_text("test\n")
        shutil.copyfile(ROOT/"tests/stream_fixture.py", scripts/"streams.py")
        (packages/"linked_module.py").write_text(
            "from xll_python import function, Reference\nVALUE = 13\n"
            "@function(name='PY.ADDRESS')\ndef address(r: Reference): return list(r.areas[0])\n")
        (scripts/"probe.py").write_text(
            "import os, sys, ssl, sqlite3, tkinter, pip, linked_module\nfrom xll_python import function\n"
            "__excel_exports__ = [linked_module.address]\n"
            "@function(name='PY.PROBE')\ndef probe(): return [os.getpid(), linked_module.VALUE, sys.prefix, pip.__version__]\n"
            "@function(name='PY.BAD')\ndef bad(): raise RuntimeError('traceback sentinel')\n")
        (profile/"settings.json").write_text(json.dumps(dict(folders=[str(ROOT/"examples"),str(scripts)],packages=[str(packages)])))
        startup = subprocess.STARTUPINFO()
        startup.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startup.wShowWindow = 0
        env = dict(os.environ,XLL_PYTHON_HOME=str(profile))
        if os.environ.get("XLL_WEB_TEST"):
            env["WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS"] = "--remote-debugging-port=0"
        proc = subprocess.Popen([excel,"/x","/automation",base/"book.csv"], env=env,startupinfo=startup)
        app = None
        try:
            app = wait_for(lambda: connect(proc.pid,base/"book.csv"))
            app.DisplayAlerts = False
            assert app.RegisterXLL(str(standalone)), "Excel rejected XLL"
            log = (profile/"python-excel.log").read_text(encoding="utf-8-sig")
            print(log,flush=True)
            assert "Loaded 10 Python worksheet functions." in log, log
            book=app.ActiveWorkbook
            sheet=book.Worksheets(1)
            sheet.Range("B1").Formula2 = "=XLLPY.scale(21)"
            assert sheet.Range("B1").Value2 == 42, sheet.Range("B1").Text
            sheet.Range("B2").Formula2 = "=XLLPY.scale(5,3)"
            assert sheet.Range("B2").Value2 == 15
            sheet.Range("E1").Formula2 = "=PY.PROBE()"
            assert sheet.Range("E1").Value2 == proc.pid, sheet.Range("E1").Text
            assert sheet.Range("E2").Value2 == 13
            home=Path(sheet.Range("E3").Value2)
            assert home.parent == profile and (home/"libs/python313.lib").exists()
            assert sheet.Range("E4").Value2
            sheet.Range("A4:B5").Value2 = ((1,2),(3,4))
            sheet.Range("G1").Formula2 = "=XLLPY.transpose(A4:B5)"
            assert sheet.Range("G1:H2").Value2 == ((1.,3.),(2.,4.)), sheet.Range("G1").Text
            sheet.Range("K1").Formula2 = "=PY.ADDRESS(A4:B5)"
            assert sheet.Range("K1:K4").Value2 == ((3.,),(4.,),(0.,),(1.,)), sheet.Range("K1").Text
            sheet.Range("C1").Formula2 = "=XLLPY.delayed(123,3)"
            assert sheet.Range("C1").Value2 == "Waiting…", sheet.Range("C1").Text
            sheet.Range("D1").Formula2 = "=XLLPY.scale(4)"
            assert sheet.Range("D1").Value2 == 8 and sheet.Range("C1").Value2 == "Waiting…"
            sheet.Range("D2").Formula2 = "=IF(ISNUMBER(C1),C1+1,0)"
            wait_for(lambda: sheet.Range("C1").Value2 == 123)
            assert sheet.Range("D2").Value2 == 124
            sheet.Range("J1").Formula2 = "=PY.BAD()"
            assert sheet.Range("J1").Text == "#VALUE!"
            assert "traceback sentinel" in (profile/"python-excel.log").read_text(encoding="utf-8-sig")
            another=app.Workbooks.Add()
            another.Worksheets(1).Range("A1").Formula2 = "=XLLPY.scale(7)"
            assert another.Worksheets(1).Range("A1").Value2 == 14
            another.Close(False)
            verify_streams(app, sheet, scripts)
            print("PASS: fresh single-XLL load, native CPython in EXCEL.EXE, pip, Tk, SSL, SQLite, package link, two workbooks, defaults, spills, nonblocking async and dependent recalculation.",flush=True)
            if "--ui" in sys.argv:
                verify(app,proc.pid,base,home)
        finally:
            if app is not None:
                def close_debugger(hwnd, _):
                    if win32process.GetWindowThreadProcessId(hwnd)[1]==proc.pid and win32gui.GetWindowText(hwnd).startswith("Python IDE"):
                        win32gui.PostMessage(hwnd,win32con.WM_CLOSE,0,0)
                win32gui.EnumWindows(close_debugger,None)
                try:
                    app.DisplayAlerts=False
                    app.Quit()
                except pythoncom.com_error:
                    pass
                app=None
            try: proc.wait(timeout=8)
            except subprocess.TimeoutExpired:
                proc.kill()
                proc.wait()


if __name__ == "__main__":
    main()

