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
import threading

sys.coinit_flags = 2
import psutil
import pythoncom
import win32con
import win32gui
import win32process
from pywinauto import Desktop
from excel_smoke import ROOT, connect, wait_for
from compiler_designer import choose_file
from web_editor import page


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def open_tools(main, pid, script):
    def invoke(label):
        pythoncom.CoInitialize()
        try:
            main.child_window(title=label,control_type="Button").invoke()
        finally:
            pythoncom.CoUninitialize()

    def window(title):
        dialog = Desktop(backend="win32").window(process=pid, title_re=title)
        dialog.wait("visible", timeout=45)
        return dialog.handle

    main.child_window(title="Python",control_type="TabItem").select()
    threading.Thread(target=invoke, args=("Run Script",), daemon=True).start()
    choose_file(pid, script)
    handles = [window("Python output.*probe.py")]
    threading.Thread(target=invoke, args=("Run Macro",), daemon=True).start()
    picker = Desktop(backend="uia").window(handle=window("Run Macro"))
    picker.child_window(title="Window test",control_type="ListItem").select()
    picker.child_window(title="Run",control_type="Button").invoke()
    handles.append(window("Python output.*Window test"))
    for label, title in (("Console","Python Console"), ("Manage Packages","Manage Packages"), ("Open IDE","Python IDE.*")):
        invoke(label)
        handles.append(window(title))
    invoke("Compile to XLL")
    compiler = wait_for(lambda:next((p for p in psutil.Process(pid).children() if p.name()=="Designer.exe"),None))
    return handles, compiler


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
        profile = base/"profile"
        profile.mkdir()
        scripts = base/"scripts"
        scripts.mkdir()
        script = scripts/"probe.py"
        script.write_text("from xll_python import macro\n@macro(name='Window test')\ndef probe(): print(42)\n")
        (profile/"settings.json").write_text(json.dumps(dict(folders=[str(scripts)],packages=[])))
        proc = subprocess.Popen([excel,"/x","/automation",csv],env=dict(os.environ,XLL_PYTHON_HOME=str(profile),
                                WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS="--remote-debugging-port=0"))
        compiler = None
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
            handles, compiler = open_tools(main, proc.pid, script)
            with page(profile/"ConsoleWebView2","https://python-console.example/") as (js, send):
                wait_for(lambda:js("!!document.querySelector('#source')"))
                js("source.focus()")
                send("Input.insertText", text="from xll_python import Application\nbook_ref = Application.ActiveWorkbook\nprint(book_ref.Name)")
                js("run.click()")
                wait_for(lambda:js("!run.disabled && transcript.innerText.includes(" + json.dumps(unsaved.Name) + ")"))
            def click():
                main.child_window(title="Python",control_type="TabItem").select()
                main.child_window(title="Check for Updates",control_type="Button").invoke()
                progress.wait("visible", timeout=3)
            click()
            press("Cancel")
            wait_for(lambda:not list(base.glob("*.tmp")))
            assert all(win32gui.IsWindow(h) for h in handles) and compiler.is_running()
            assert app.Workbooks.Count == 3 and not app.Workbooks(files[0].name).Saved
            assert unsaved.Worksheets(1).Range("A1").Value2=="unsaved text" and digest(target)==before
            print("PASS: declining restart left workbooks open, unsaved changes intact and XLL unchanged",flush=True)
            if action == "Save all":
                click()
                press(action)
                win32gui.PostMessage(save_dialog(),win32con.WM_CLOSE,0,0)
                wait_for(lambda:not list(base.glob("*.tmp")))
                assert app.Workbooks.Count == 3
                assert all(win32gui.IsWindow(h) for h in handles) and compiler.is_running()
                assert unsaved.Worksheets(1).Range("A1").Value2=="unsaved text" and digest(target)==before
                print("PASS: canceling Save As kept the untitled workbook and its edits open",flush=True)
                app.Workbooks(files[0].name).Worksheets(1).Range("A1").Value2="last edit"
            else:
                with page(profile/"CompilerWebView2", "https://xll-designer.example/") as (js, send):
                    for event in ("dragEnter", "dragOver", "drop"):
                        send("Input.dispatchDragEvent", type=event, x=140, y=160,
                             data=dict(items=[], files=[str(scripts)], dragOperationsMask=1))
                    wait_for(lambda:js("!!project && !busy"))
                    click()
                    press(action)
                    wait_for(lambda:js("$('leave').open"))
                    js("$('cancelLeave').click()")
                    wait_for(lambda:not list(base.glob("*.tmp")))
                    assert all(win32gui.IsWindow(h) for h in handles) and compiler.is_running()
                    assert unsaved.Worksheets(1).Range("A1").Value2=="unsaved text" and digest(target)==before
                    print("PASS: canceling the project close preserved the tools and unsaved workbook",flush=True)
                    click()
                    press(action)
                    wait_for(lambda:js("$('leave').open"))
                    prompt = Desktop(backend="win32").window(process=proc.pid,title="Python for Excel update")
                    stop = next(w for w in prompt.descendants() if w.window_text()=="Cancel update")
                    win32gui.PostMessage(stop.handle,win32con.BM_CLICK,0,0)
                    wait_for(lambda:not list(base.glob("*.tmp")))
                    js("$('discard').click()")
                    wait_for(lambda:js("!busy"))
                    assert compiler.is_running() and app.Workbooks.Count == 3
                    print("PASS: canceling the update withdrew the compiler close request",flush=True)
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
            else:
                with page(profile/"CompilerWebView2", "https://xll-designer.example/") as (js, _):
                    wait_for(lambda:js("$('leave').open"))
                    designer = Desktop(backend="win32").window(process=compiler.pid,title="Compile to XLL")
                    win32gui.SendMessage(designer.handle,win32con.WM_CLOSE,0,0)
                    js("$('discard').click()")
            exit_code = proc.wait(100)
            new_pid=wait_for(replacement_pid,35)
            assert exit_code == 0, f"Excel exited with {exit_code:#x}"
            assert not compiler.is_running() and not any(win32gui.IsWindow(h) for h in handles)
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
            if compiler and compiler.is_running():
                compiler.kill()
                compiler.wait()
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
