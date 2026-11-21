"""Exercise save, startup-folder prompts, and workspace persistence in desktop Excel."""
import json
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
import time

import win32con
import win32gui
from pywinauto import Desktop

from excel_smoke import ROOT, connect, wait_for
from web_editor import page


def verify(base, proc, app):
    scripts, extra, scratch, packages = (base / name for name in ("scripts", "extra", "scratch", "packages"))
    profile = base / "profile"
    settings_file = profile / "settings.json"
    desktop = Desktop(backend="win32")
    main = Desktop(backend="uia").window(handle=app.Hwnd)
    main.child_window(title="Python", control_type="TabItem").select()

    def open_ide():
        main.child_window(title="Open IDE", control_type="Button").invoke()
        return page(profile / "IdeWebView2", "http://xll-")

    def close_ide():
        window = desktop.window(process=proc.pid, title_re="Python IDE.*")
        win32gui.PostMessage(window.handle, win32con.WM_CLOSE, 0, 0)
        wait_for(lambda: not window.exists(timeout=.1))

    def settings():
        return json.loads(settings_file.read_text(encoding="utf-8-sig"))

    def answer(text, folder, yes):
        dialog = desktop.window(process=proc.pid, title="Excel startup folders")
        dialog.wait("visible", timeout=12)
        message = " ".join(dialog.texts() + [item.window_text() for item in dialog.descendants()])
        assert text in message and str(folder).casefold() in message.casefold(), message
        win32gui.PostMessage(win32gui.GetDlgItem(dialog.handle, 6 if yes else 7), win32con.BM_CLICK, 0, 0)
        wait_for(lambda: not dialog.exists(timeout=.1))

    def ready(js):
        js("[...document.querySelectorAll('button')].find(b => b.innerText.includes('Yes, I trust'))?.click()")
        return f"Excel {proc.pid}" in js("document.body.innerText").replace("\xa0", " ")

    cell = app.ActiveSheet.Range("A2")
    cell.Formula2 = "=PY.SAVED()"
    assert cell.Value2 == 10
    with open_ide() as (js, send):
        wait_for(lambda: ready(js))
        send("Page.bringToFront")

        def key(name, vk, modifiers=0):
            for event in ("keyDown", "keyUp"):
                send("Input.dispatchKeyEvent", type=event, key=name, windowsVirtualKeyCode=vk,
                     nativeVirtualKeyCode=vk, modifiers=modifiers)

        def open_file(name):
            key("p", 80, 2)
            wait_for(lambda: js("!!document.querySelector('.quick-input-widget input:focus')"))
            send("Input.insertText", text=name)
            wait_for(lambda: name == js("document.querySelector('.quick-input-list .monaco-list-row.focused .label-name')?.textContent"))
            key("Enter", 13)
            wait_for(lambda: js("document.querySelector('.monaco-editor.focused')?.getAttribute('data-uri') || ''").endswith('/' + name))

        def replace_text(text):
            key("a", 65, 2)
            send("Input.insertText", text=text)

        def save(text):
            replace_text(text)
            key("s", 83, 2)

        open_file("_factor.py")
        save("VALUE = 20\n")
        wait_for(lambda: cell.Value2 == 20)
        save("VALUE = * 10\n")
        wait_for(lambda: "Previous functions remain active" in js("document.body.innerText").replace("\xa0", " "))
        assert cell.Value2 == 20
        save("VALUE = 30\n")
        wait_for(lambda: cell.Value2 == 30)
        print("PASS: Ctrl+S updated a worksheet through a nested helper and recovered from a syntax error.", flush=True)

        def command(label):
            key("F1", 112)
            wait_for(lambda: js("!!document.querySelector('.quick-input-widget input:focus')"))
            send("Input.insertText", text=label)
            wait_for(lambda: label in js("document.querySelector('.quick-input-list')?.innerText || ''"))
            key("Enter", 13)

        def refresh_count():
            return (profile / "python-excel.log").read_text(encoding="utf-8-sig").count("Loaded 1 Python worksheet functions.")

        count = refresh_count()
        open_file("_offset.py")
        replace_text("OFFSET = 3\n")
        open_file("_factor.py")
        replace_text("VALUE = 31\n")
        command("Save All")
        wait_for(lambda: cell.Value2 == 34)
        time.sleep(.6)
        assert refresh_count() == count + 1
        print("PASS: Save All reloaded both edits and recalculated once.", flush=True)

        def change(folder, add):
            def click_row(container):
                row = f"[...document.querySelectorAll('{container} .theia-TreeNode')].find(e => e.title.toLowerCase() === {json.dumps(str(folder).lower())})"
                wait_for(lambda: js(f"!!({row})"), 10)
                x, y = js(f"(() => {{ const r = ({row}).getBoundingClientRect(); return [r.x + 45, r.y + 8] }})()")
                for event in ("mousePressed", "mouseReleased"):
                    send("Input.dispatchMouseEvent", type=event, x=x, y=y, button="left", clickCount=1)
            if add:
                command("Add Folder to Workspace")
                selector = "document.querySelector('.dialogContent select.theia-LocationList')"
                wait_for(lambda: js(f"Array.from({selector}?.options || []).some(o => o.text === {json.dumps(base.name)})"))
                js(f"""var list = {selector};
                    list.value = [...list.options].find(o => o.text === {json.dumps(base.name)}).value;
                    list.dispatchEvent(new Event('change', {{bubbles:true}}))""")
                click_row(".dialogContent")
                action = "Open"
            else:
                root = f"[...document.querySelectorAll('#files .theia-TreeNode')].find(n => n.title.toLowerCase() === {json.dumps(str(folder).lower())})"
                js(f"""var root = ({root}); document.querySelector('#files').focus(); root.click();
                    var r = root.getBoundingClientRect();
                    root.dispatchEvent(new MouseEvent('contextmenu', {{bubbles:true, cancelable:true, button:2,
                        clientX:r.x+30, clientY:r.y+8}}))""")
                menu = """document.querySelector("[data-command='workspace:removeFolder']")"""
                wait_for(lambda: js(f"!!({menu})"), 8)
                for _ in range(30):
                    key("ArrowDown", 40)
                    if js(f"({menu}).classList.contains('lm-mod-active')"):
                        break
                else:
                    raise AssertionError("Remove Folder from Workspace was unavailable")
                key("Enter", 13)
                action = "OK"
            button = f"[...document.querySelectorAll('.dialogControl button')].find(b => b.innerText.toLowerCase() === {json.dumps(action.lower())})"
            wait_for(lambda: js(f"!!({button})"))
            js(f"({button}).click()")

        change(scratch, True)
        answer("load worksheet functions", scratch, False)
        wait_for(lambda: js("[...document.querySelectorAll('button')].some(b => b.innerText.includes('Yes, I trust'))"))
        ready(js)
        assert scratch not in map(Path, settings()["folders"])
        before = refresh_count()
        open_file("scratch_only.py")
        save("VALUE = 100\n")
        wait_for(lambda: "100" in (scratch / "scratch_only.py").read_text())
        time.sleep(.6)
        assert refresh_count() == before
        change(extra, True)
        answer("load worksheet functions", extra, True)
        wait_for(lambda: extra in map(Path, settings()["folders"]))
        wait_for(lambda: js("[...document.querySelectorAll('button')].some(b => b.innerText.includes('Yes, I trust'))"))
        ready(js)
        change(extra, False)
        answer("remove this worksheet folder", extra, False)
        assert extra in map(Path, settings()["folders"])
        change(packages, False)
        answer("package links", packages, True)
        wait_for(lambda: settings()["packages"] == [])
        change(scripts, False)
        answer("remove this worksheet folder", scripts, True)
        wait_for(lambda: list(map(Path, settings()["folders"])) == [extra])
    close_ide()

    # Exercise the same persisted changes made by a ribbon add, then removal, of the Explorer-only folder.
    for configured in ([extra, scratch], [extra]):
        selected = settings()
        selected["folders"] = list(map(str, configured))
        settings_file.write_text(json.dumps(selected), encoding="utf-8")
        with open_ide() as (js, _):
            wait_for(lambda: ready(js))
            wait_for(lambda: scratch.name in js("document.querySelector('#files')?.innerText || ''"))
            roots = js("Array.from(document.querySelectorAll('#files .theia-DirNode'), e => e.title.toLowerCase())")
            assert str(scratch).lower() in roots and not any(str(p).lower() in roots for p in (scripts, extra, packages)), roots
        close_ide()
    print("PASS: accepted and declined startup changes, package-link removal, and independent roots after reopen.", flush=True)


def main():
    excel = Path(os.environ["ProgramFiles"]) / "Microsoft Office/root/Office16/EXCEL.EXE"
    with tempfile.TemporaryDirectory(dir=ROOT / ".build") as tmp:
        base = Path(tmp)
        for name in ("scripts/helpers", "extra", "scratch", "packages", "profile"):
            (base / name).mkdir(parents=True)
        (base / "profile/settings.json").write_text(json.dumps({
            "folders": [str(base / "scripts")], "packages": [str(base / "packages")]}))
        (base / "scripts/helpers/_factor.py").write_text("VALUE = 10\n")
        (base / "scripts/_offset.py").write_text("OFFSET = 0\n")
        (base / "scratch/scratch_only.py").write_text("VALUE = 0\n")
        (base / "scripts/saved.py").write_text(
            "from xll_python import function\nfrom helpers._factor import VALUE\nfrom _offset import OFFSET\n"
            "@function(name='PY.SAVED')\ndef saved(): return VALUE + OFFSET\n")
        book = base / "book.csv"
        book.write_text("test\n")
        xll = base / "xll-python.xll"
        shutil.copyfile(ROOT / "dist/xll-python.xll", xll)
        proc = subprocess.Popen([excel, "/x", "/automation", book], env=dict(os.environ,
            XLL_PYTHON_HOME=str(base / "profile"), WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS="--remote-debugging-port=0"))
        app = None
        try:
            app = wait_for(lambda: connect(proc.pid, book))
            app.DisplayAlerts = False
            assert app.RegisterXLL(str(xll))
            app.Visible = True
            verify(base, proc, app)
        finally:
            if app:
                app.Quit()
            try:
                proc.wait(timeout=15)
            except subprocess.TimeoutExpired:
                proc.kill()
                proc.wait()


if __name__ == "__main__":
    main()
