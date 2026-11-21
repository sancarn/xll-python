"""Optional interactive smoke test in the smoke driver's disposable Excel process."""
import json
import os
from pathlib import Path
import threading
import time
import zipfile

import pythoncom
import win32con
import win32gui
import win32process
from pywinauto import Desktop
from pywinauto.controls.uiawrapper import UIAWrapper
from excel_smoke import wait_for


def fixture_wheel(base, name):
    wheel = base / f"{name}-1.0-py3-none-any.whl"
    info = f"{name}-1.0.dist-info/"
    with zipfile.ZipFile(wheel, "w") as z:
        files = {name + ".py": "VALUE = 42\n",
                 info + "METADATA": f"Metadata-Version: 2.1\nName: {name}\nVersion: 1.0\n",
                 info + "WHEEL": "Wheel-Version: 1.0\nRoot-Is-Purelib: true\nTag: py3-none-any\n"}
        files[info + "RECORD"] = "".join(f"{k},,\n" for k in [*files, info + "RECORD"])
        for path, source in files.items():
            z.writestr(path, source)
    return wheel


def verify(app, pid, base, home):
    app.Visible = True
    desktop = Desktop(backend="uia")
    main = desktop.window(handle=app.Hwnd)
    tab = main.child_window(title="Python", control_type="TabItem").wrapper_object()
    tab.select()
    def button(label):
        return main.child_window(title=label,control_type="Button").wrapper_object()
    def click_ribbon(label):
        # Modal dialogs can keep UIA Invoke on the calling thread until dismissed.
        def invoke():
            pythoncom.CoInitialize()
            try: button(label).invoke()
            finally: pythoncom.CoUninitialize()
        threading.Thread(target=invoke,daemon=True).start()
    def hwnd_for(title="", cls=None,children=False):
        pids={pid}
        if children:
            import psutil
            pids.update(p.pid for p in psutil.Process(pid).children(recursive=True))
        found=[]
        def visit(h,_):
            if win32process.GetWindowThreadProcessId(h)[1] in pids and win32gui.IsWindowVisible(h):
                if (cls is None or win32gui.GetClassName(h)==cls) and win32gui.GetWindowText(h).startswith(title):
                    found.append(h)
        win32gui.EnumWindows(visit,None)
        return found[0] if found else 0
    def window(title):
        return desktop.window(handle=wait_for(lambda:hwnd_for(title),20))
    def close(w):
        win32gui.PostMessage(w.handle,win32con.WM_CLOSE,0,0)
    def pick_file(path,children=False):
        handle=wait_for(lambda:hwnd_for(cls="#32770",children=children),12)
        edits=[]
        win32gui.EnumChildWindows(handle,lambda h,_:edits.append(h) if win32gui.GetClassName(h)=="Edit" else None,None)
        assert edits, "Missing file name edit"
        filename=next((h for h in edits if win32gui.GetDlgCtrlID(h)==1148),edits[0])
        win32gui.SendMessage(filename,win32con.WM_SETTEXT,0,str(path))
        win32gui.PostMessage(win32gui.GetDlgItem(handle,1),win32con.BM_CLICK,0,0)
    def read_console(w):
        edits=[e for e in w.descendants(control_type="Edit") if e.element_info.class_name.startswith("WindowsForms10.EDIT")]
        return max((e.get_value() for e in edits),key=len,default="")
    for label in ("Run Script","Manage Packages","Startup Folders","Open IDE","Check for Updates","Console"):
        button(label)
    print("PASS: Python ribbon tab and six actionable buttons",flush=True)
    click_ribbon("Console")
    console = window("Python Console")
    if os.environ.get("XLL_WEB_TEST"):
        from web_editor import verify_console, verify_reopened_console
        verify_console(base, pid, app.ActiveSheet)
    close(console)
    if os.environ.get("XLL_WEB_TEST"):
        wait_for(lambda: not hwnd_for("Python Console"))
        click_ribbon("Console")
        console = window("Python Console")
        verify_reopened_console(base)
        wait_for(lambda: app.ActiveSheet.Range("N1").Text in ("#NAME?", "#VALUE!"))
        close(console)

    click_ribbon("Startup Folders")
    folders=window("Python startup folders")
    folders.child_window(title="Package links",control_type="TabItem").select()
    assert str(base/"packages") in [v.window_text() for v in folders.descendants(control_type="ListItem")]
    folders.child_window(title="Save",control_type="Button").invoke()
    print("PASS: startup folder / package link settings round-trip",flush=True)

    def refresh():
        click_ribbon("Startup Folders")
        window("Python startup folders").child_window(title="Refresh",control_type="Button").invoke()

    sheet=app.ActiveSheet
    helper=base/"scripts/_rates.py"
    helper.write_text("RATE = 11\n")
    script=base/"scripts/reload.py"
    live_rate = (
        "import asyncio\nfrom pathlib import Path\n"
        "@function(name='PY.LIVE_RATE')\nasync def live_rate():\n"
        "    try:\n        yield RATE\n        await asyncio.Event().wait()\n"
        "    finally:\n        Path(__file__).with_name(f'closed-rate-{RATE}').touch()\n"
    )
    script.write_text("from xll_python import function\nfrom _rates import RATE\n"
                      "@function(name='PY.RATE')\ndef rate(): return RATE + 1\n"
                      "@function(name='PY.OLD')\ndef old(): return 7\n" + live_rate)
    refresh()
    wait_for(lambda:app.Evaluate("=PY.RATE()")==12)
    sheet.Range("M1").Formula2="=PY.RATE()"
    sheet.Range("M2").Formula2="=PY.OLD()"
    sheet.Range("M3").Formula2="=PY.LIVE_RATE()"
    wait_for(lambda:sheet.Range("M3").Value2==11)
    assert sheet.Range("M2").Value2==7
    # Same byte length and timestamp must still reload a previously imported helper.
    stamp=helper.stat()
    helper.write_text("RATE = 22\n")
    os.utime(helper,ns=(stamp.st_atime_ns,stamp.st_mtime_ns))
    script.write_text("from xll_python import function\nfrom _rates import RATE\n"
                      "@function(name='PY.RATE')\ndef rate(extra=3): return RATE + extra\n"
                      "@function(name='PY.NEW')\ndef new(): return 8\n" + live_rate)
    refresh()
    wait_for(lambda:sheet.Range("M1").Value2==25)
    wait_for(lambda:sheet.Range("M3").Value2==22)
    wait_for((script.parent/"closed-rate-11").exists)
    assert app.Evaluate("=PY.RATE(5)")==27 and app.Evaluate("=PY.NEW()")==8
    assert sheet.Range("M2").Text in ("#NAME?", "#VALUE!"), sheet.Range("M2").Text
    print("PASS: live refresh reloaded source and imports, updated signatures, added and removed functions, and recalculated existing cells",flush=True)

    wheel=fixture_wheel(base, "xll_probe")
    click_ribbon("Manage Packages")
    installer=window("Manage Packages")
    native_wheel=os.environ.get("XLL_TEST_NATIVE_WHEEL")
    install=f'install --no-index "{wheel}"'
    if native_wheel:
        install+=' "'+str(Path(native_wheel).resolve(strict=True))+'"'
    if os.environ.get("XLL_WEB_TEST"):
        from web_editor import verify_packages, verify_package_log
        for name in ("openpyxl", "xlsxwriter"):
            fixture_wheel(base, name)
        verify_packages(base, install)
        close(installer)
        wait_for(lambda: not hwnd_for("Manage Packages"))
        click_ribbon("Manage Packages")
        installer=window("Manage Packages")
        verify_package_log(base)
    else:
        installer.child_window(title="Advanced",control_type="TabItem").select()
        installer.child_window(title="pip",control_type="Edit").set_edit_text(install)
        installer.child_window(title="Run",control_type="Button").invoke()
        wait_for(lambda: installer.child_window(title="Complete · exit 0",control_type="Text").exists(timeout=.1))
    close(installer)
    print("PASS: Manage Packages installed a wheel offline",flush=True)

    script=base/"interactive.py"
    script.write_text("from xll_probe import VALUE\namount = VALUE + 1\nprint('FROM_XLL', amount)\n",encoding="utf-8")
    if native_wheel:
        with script.open("a") as f:
            f.write("from markupsafe import _speedups\nassert _speedups.__file__.endswith('.pyd')\nprint('NATIVE_WHEEL_LOADED')\n")
    click_ribbon("Run Script")
    pick_file(script)
    output=window("Python output")
    wait_for(lambda:"FROM_XLL 43" in read_console(output))
    if native_wheel:
        assert "NATIVE_WHEEL_LOADED" in read_console(output)
    close(output)
    print("PASS: Run Script picked a local file and imported the installed package",flush=True)

    click_ribbon("Open IDE")
    ide=window("Python IDE")
    # Use the separately named native target selector for files opened through workspace folders.
    def choose():
        pythoncom.CoInitialize()
        try: ide.child_window(title="Choose script…",control_type="Button").invoke()
        finally: pythoncom.CoUninitialize()
    threading.Thread(target=choose,daemon=True).start()
    pick_file(script)
    def action(label):
        ide.child_window(title=label,control_type="Button").invoke()
    action("Arm breakpoint")
    action("Run saved script")
    wait_for(lambda:"interactive.py" in read_console(ide) and "line 1 " in read_console(ide))
    action("Step")
    wait_for(lambda:"line 2 " in read_console(ide))
    action("Continue")
    wait_for(lambda:"FROM_XLL 43" in read_console(ide))
    edit=next(e for e in ide.descendants(control_type="Edit") if not e.iface_value.CurrentIsReadOnly
              and e.element_info.class_name.startswith("WindowsForms10.EDIT"))
    edit.set_edit_text("=XLLPY.scale(11)")
    action("Evaluate")
    wait_for(lambda:"=XLLPY.scale(11) = 22" in read_console(ide))
    print("PASS: IDE breakpoint, locals, step, resume, run and Excel formula evaluation",flush=True)
    threading.Thread(target=choose,daemon=True).start()
    pick_file(Path(__file__).resolve().parents[1]/"examples/functions.py")
    action("Arm breakpoint")
    edit.set_edit_text("=XLLPY.scale(19)")
    action("Evaluate")
    wait_for(lambda:"scale line" in read_console(ide) and '"value":"19.0"' in read_console(ide))
    action("Continue")
    wait_for(lambda:"=XLLPY.scale(19) = 38" in read_console(ide))
    print("PASS: paused a live registered worksheet function, inspected its arguments and returned its result to Excel",flush=True)
    if os.environ.get("XLL_WEB_TEST"):
        from web_editor import verify_editor
        verify_editor(ide,script,read_console,action,pick_file)
    close(ide)

