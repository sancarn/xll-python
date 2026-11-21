"""Interactive updater regression. Pass an XLL built with PYTHONEXCEL_RELEASE_TAG=v0.0.0."""
import gc
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile

sys.coinit_flags = 2
import psutil
import pythoncom
import win32con
import win32gui
import win32process
from pywinauto import Desktop
from excel_smoke import ROOT, connect, wait_for


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def test(source, action):
    url = "https://api.github.com/repos/sancarn/xll-python/releases/latest"
    release = json.loads(subprocess.check_output(["curl.exe", "--fail", "--silent", "--show-error", url], timeout=30))
    assets = {a["name"]: a for a in release["assets"]}
    asset = assets.get("xll-python.xll") or assets["PythonExcel.xll"]
    sha = asset["digest"].removeprefix("sha256:")
    excel = Path(os.environ["ProgramFiles"]) / "Microsoft Office/root/Office16/EXCEL.EXE"
    with tempfile.TemporaryDirectory(dir=ROOT/".build",prefix="restart-") as temp:
        base = Path(temp)
        target = base / ("xll-python.xll" if action == "Save all" else "Python Excel.xll")
        shutil.copyfile(source, target)
        before = digest(target)
        assert sha != before
        csv = base/"startup.csv"
        csv.write_text("test\n")
        files = [base/"One workbook.xlsx",base/"Two \u03a9.xlsx"]
        proc = subprocess.Popen([excel,"/x","/automation",csv],env=dict(os.environ,XLL_PYTHON_HOME=str(base/"profile")))
        app = restored = None
        progress = Desktop(backend="win32").window(process=proc.pid, title="Checking for updates")
        def dialogs(pid):
            windows=[]
            def visit(h,_):
                if (win32process.GetWindowThreadProcessId(h)[1]==pid and win32gui.IsWindowVisible(h)
                        and win32gui.GetClassName(h)=="#32770"):
                    windows.append(h)
            win32gui.EnumWindows(visit,None)
            return windows
        def press(label):
            def find():
                prompt = Desktop(backend="win32").window(process=proc.pid, title="Python for Excel update")
                if prompt.exists(timeout=0):
                    assert not progress.exists(timeout=0), "Spinner still visible at the workbook prompt"
                    buttons = {w.window_text(): w for w in prompt.descendants()}
                    assert {"Save all", "Don't save any", "Cancel"} <= buttons.keys(), list(buttons)
                    win32gui.PostMessage(buttons[label].handle,win32con.BM_CLICK,0,0)
                    return True
            wait_for(find,40)
        def save_dialog():
            return wait_for(lambda:next((h for h in dialogs(proc.pid) if win32gui.GetWindowText(h)=="Save As"),None))
        def replacement_pid():
            for p in psutil.process_iter(["pid","name","cmdline"]):
                if (p.info["name"] or "").lower()=="excel.exe" and str(files[0]) in (p.info["cmdline"] or []):
                    return p.pid
        try:
            app=wait_for(lambda:connect(proc.pid,csv))
            app.Visible=True
            assert app.RegisterXLL(str(target))
            for file in files:
                book=app.Workbooks.Add()
                book.Worksheets(1).Range("A1").Value2="original"
                book.SaveAs(str(file),51)
                book.Worksheets(1).Range("A1").Value2="saved by updater"
                del book
            app.Workbooks(csv.name).Close(False)
            unsaved=app.Workbooks.Add()
            unsaved.Worksheets(1).Range("A1").Value2="unsaved text"
            main=Desktop(backend="uia").window(handle=app.Hwnd)
            def click():
                main.child_window(title="Python",control_type="TabItem").select()
                main.child_window(title="Check for Updates",control_type="Button").invoke()
                progress.wait("visible", timeout=3)
            click()
            press("Cancel")
            wait_for(lambda:not list(base.glob("*.tmp")))
            assert app.Workbooks.Count == 3 and not app.Workbooks(files[0].name).Saved
            assert unsaved.Worksheets(1).Range("A1").Value2=="unsaved text" and digest(target)==before
            print("PASS: declining restart left workbooks open, unsaved changes intact and XLL unchanged",flush=True)
            if action == "Save all":
                click()
                press(action)
                win32gui.PostMessage(save_dialog(),win32con.WM_CLOSE,0,0)
                wait_for(lambda:not list(base.glob("*.tmp")))
                assert app.Workbooks.Count == 3
                assert unsaved.Worksheets(1).Range("A1").Value2=="unsaved text" and digest(target)==before
                print("PASS: canceling Save As kept the untitled workbook and its edits open",flush=True)
                app.Workbooks(files[0].name).Worksheets(1).Range("A1").Value2="last edit"
            del unsaved
            click()
            app=None
            gc.collect()
            press(action)
            new_file = base/"New workbook.xlsx"
            if action == "Save all":
                dialog = Desktop(backend="win32").window(handle=save_dialog())
                dialog.child_window(class_name="Edit", found_index=0).set_edit_text(str(new_file))
                win32gui.PostMessage(dialog.handle,win32con.WM_COMMAND,win32con.IDOK,0)
            proc.wait(100)
            new_pid=wait_for(replacement_pid,35)
            def reconnect():
                for h in Desktop(backend="win32").windows(process=new_pid,visible_only=True):
                    if h.window_text()=="Microsoft Excel Security Notice":
                        notice=Desktop(backend="uia").window(handle=h.handle)
                        paths=[Path(e.window_text()) for e in notice.descendants(control_type="Edit")]
                        if target in paths and digest(target)==sha:
                            notice.child_window(title="Enable this add-in for this session only.",control_type="Button").invoke()
                return connect(new_pid,files[0])
            restored=wait_for(reconnect,150)
            expected = ["last edit", "saved by updater"] if action == "Save all" else ["original", "original"]
            for file, value in zip(files, expected):
                assert restored.Workbooks(file.name).Worksheets(1).Range("A1").Value2==value
            assert restored.Workbooks.Count == (3 if action == "Save all" else 2)
            if action == "Save all":
                assert restored.Workbooks(new_file.name).Worksheets(1).Range("A1").Value2=="unsaved text"
            assert digest(target)==sha
            print(f"PASS: {action} restarted Excel with the expected workbooks and contents",flush=True)
        finally:
            if restored is not None:
                restored.DisplayAlerts=False
                restored.Quit()
            new_pid=replacement_pid()
            if new_pid:
                p=psutil.Process(new_pid)
                try:p.wait(8)
                except psutil.TimeoutExpired:p.kill();p.wait()
            if proc.poll() is None:
                proc.kill()
                proc.wait()


if __name__=="__main__":
    for action in ("Save all", "Don't save any"):
        test(Path(sys.argv[1]), action)
