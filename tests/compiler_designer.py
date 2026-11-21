"""Exercise the packaged designer, native dialogs, WebView2 preview and real compiler."""
import json
import subprocess

from excel_smoke import wait_for
from web_editor import page
import win32con
import win32gui
import win32process


def dialog(pid):
    found = []
    def visit(hwnd, _):
        if (win32process.GetWindowThreadProcessId(hwnd)[1] == pid and win32gui.IsWindowVisible(hwnd)
                and win32gui.GetClassName(hwnd) == "#32770"):
            found.append(hwnd)
    win32gui.EnumWindows(visit, None)
    return found[0] if found else None


def choose_file(pid, path):
    hwnd = wait_for(lambda: dialog(pid), 15)
    edits = []
    win32gui.EnumChildWindows(hwnd, lambda h, _: edits.append(h) if win32gui.GetClassName(h) == "Edit" else None, None)
    filename = next((h for h in edits if win32gui.GetDlgCtrlID(h) == 1148), edits[0])
    win32gui.SendMessage(filename, win32con.WM_SETTEXT, 0, str(path))
    win32gui.PostMessage(win32gui.GetDlgItem(hwnd, 1), win32con.BM_CLICK, 0, 0)


def drive_designer(python, support, source, template, output, env, wheel=None):
    profile = output.parent / ("web-" + output.stem)
    proc = subprocess.Popen([support/"Designer.exe", python, template, profile],
                            env=dict(env, WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS="--remote-debugging-port=0"))
    try:
        with page(profile, "https://xll-designer.example/") as (js, send):
            def idle():
                wait_for(lambda: js("typeof busy !== 'undefined' && !busy"))
                assert not js("document.querySelector('#status').classList.contains('error')"), js("$('status').innerText")
            def click(id):
                js(f"$({json.dumps(id)}).click()")
            def fill(id, value, event="input"):
                js(f"$({id!r}).value={json.dumps(value)}; $({id!r}).dispatchEvent(new Event({event!r}))")
            def section(tab):
                js(f"document.querySelector('[data-tab={tab}]').click()")
                idle()
            def picker(id, path):
                click(id)
                choose_file(proc.pid, path)
                idle()
            idle()
            assert js("$('newProject').offsetWidth && $('openProject').offsetWidth")
            if wheel:
                picker("openProject", source)
                original = json.loads(source.read_text())
                assert js("project") == original
                assert not js("$('optimise').checked")
                section("packages")
                js("$('included').querySelector('button').click()")
                picker("wheel", wheel)
                assert js("project.packages") == [wheel.as_posix()]
            else:
                # A filesystem directory dropped by Chromium is delivered as a WebView2File.
                for kind in ("dragEnter", "dragOver", "drop"):
                    send("Input.dispatchDragEvent", type=kind, x=140, y=160,
                         data=dict(items=[], files=[str(source.parent)], dragOperationsMask=1))
                wait_for(lambda: js("!!project"))
                idle()
                js("document.querySelector('button[title=\"reports.py\"]').click()")
                idle()
                assert not js("$('optimise').checked")
                click("optimise")
                assert "def refresh" in js("$('source').textContent")
                fill("role", "entry", "change")
                assets = source.parent / "assets"
                assets.mkdir()
                (assets / "example.txt").write_text("example data")
                click("refresh")
                idle()
                js("document.querySelector('button[title=\"assets/example.txt\"]').click()")
                idle()
                assert js("$('role').value") == "excluded" and js("$('role').options[0].disabled")
                fill("role", "helper", "change")
                hidden = source.parent / ".config"
                hidden.mkdir()
                (hidden / "settings.json").write_text("{}")
                # Windows pickers can return a junction path for the same project folder.
                import _winapi
                alias = output.parent / "source alias"
                _winapi.CreateJunction(str(source.parent), str(alias))
                try:
                    picker("addFiles", alias / ".config/settings.json")
                finally:
                    alias.rmdir()
                assert js("selectedFile") == ".config/settings.json"
                click("home")
                wait_for(lambda: js("$('leave').open"))
                click("cancelLeave")
                idle()
                section("packages")
                js("document.querySelector('[aria-label=\"Include xll_fixture==1\"]').click()")
                section("ribbon")
                click("addTab")
                fill("label", "First")
                click("addGroup")
                fill("label", "Reports")
                click("addButton")
                fill("label", "Refresh First")
                assert "reports.refresh" in js("Array.from($('commands').options, o=>o.value)")
                fill("command", "reports.refresh ")
                picker("importImage", source.parent/"icon.png")
                wait_for(lambda: js("document.querySelector('.control.big img')?.naturalWidth"))
                for _ in range(2):
                    click("addButton")
                    click("large")
                assert js("""(() => {
                    const [a, b] = Array.from(document.querySelectorAll('.control:not(.big)'), el => el.getBoundingClientRect());
                    return a.bottom <= b.top;
                })()""")
                click("removeNode")
                js("document.querySelector('.control:not(.big)').click()")
                click("removeNode")
                manifest = source.parent / "Designer.xllproject.json"
                picker("save", manifest)
                saved = json.loads(manifest.read_text(encoding="utf-8"))
                assert saved["entries"] == ["reports.py"] and set(saved["files"]) == {"helper.py", "assets/example.txt", ".config/settings.json"}
                assert saved["packages"] == ["xll_fixture==1"]
                click("home")
                idle()
                picker("openProject", manifest)
                assert js("project") == saved
                assert saved["optimise"] and js("$('optimise').checked")
            click("compile")
            choose_file(proc.pid, output)
            wait_for(lambda: js("!busy"), 90)
            idle()
            assert output.is_file()
            js("window.chrome.webview.postMessage({id:999,action:'close'})")
        proc.wait(15)
    finally:
        if proc.poll() is None:
            proc.kill()
            proc.wait()
