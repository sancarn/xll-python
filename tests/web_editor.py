"""Exercise the add-in's WebView2 pages using the smoke driver's disposable profiles."""
import json
from contextlib import contextmanager
import urllib.request

import websocket
from excel_smoke import wait_for


@contextmanager
def page(profile, url):
    def target():
        path = next(profile.rglob("DevToolsActivePort"), None)
        if path:
            try:
                port = int(path.read_text().splitlines()[0])
                with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/list",timeout=3) as r:
                    return next((t for t in json.load(r) if t["type"]=="page" and t["url"].startswith(url)),None)
            except (OSError, ValueError, IndexError):
                return None
    endpoint=wait_for(target)["webSocketDebuggerUrl"]
    ws=websocket.create_connection(endpoint,suppress_origin=True,timeout=10)
    seq=0
    def send(method,**params):
        nonlocal seq
        seq+=1
        ws.send(json.dumps(dict(id=seq,method=method,params=params)))
        while True:
            msg=json.loads(ws.recv())
            if msg.get("id")==seq:
                if "error" in msg:
                    raise RuntimeError(msg["error"])
                return msg["result"]
    def js(expression):
        result=send("Runtime.evaluate",expression=expression,returnByValue=True)
        if "exceptionDetails" in result:
            raise RuntimeError(result["exceptionDetails"])
        return result["result"].get("value")
    try:
        yield js, send
    finally:
        ws.close()


def verify_console(base, pid, sheet):
    with page(base/"profile"/"ConsoleWebView2","https://python-console.example/") as (js, send):
        wait_for(lambda: js("!!document.querySelector('#source')"))
        def type_code(code):
            js("source.focus(); source.select()")
            send("Input.insertText",text=code)
        def run_code(code):
            type_code(code)
            js("run.click()")
            wait_for(lambda: js("!run.disabled"))
        def rendered(selector):
            return js(f"document.querySelector({json.dumps(selector)})?.textContent ?? ''")
        run_code("import os, sys, time\nrecord = {'values': [os.getpid()]}\nprint(record)")
        js("document.querySelector('.inspect').click()")
        wait_for(lambda: rendered(".children .inspect"))
        js("document.querySelector('.children .inspect').click()")
        wait_for(lambda: str(pid) in rendered(".children .children"))
        type_code("print('first', flush=True)\ntime.sleep(1.5)\nprint('second', file=sys.stderr)\nrecord")
        assert js("highlight.querySelectorAll('.str').length") == 2
        js("run.click()")
        wait_for(lambda: js("run.disabled && transcript.lastChild.querySelector('.output').innerText.includes('first')"))
        wait_for(lambda: js("!run.disabled"))
        assert "second" in rendered(".entry:last-child .error")
        assert "dict" in rendered(".entry:last-child .result")
        type_code("record.it")
        wait_for(lambda: rendered("#suggestions [role=option]") == "items")
        send("Input.dispatchKeyEvent", type="keyDown", key="Tab", code="Tab", windowsVirtualKeyCode=9)
        assert js("source.value") == "record.items"
        run_code("if True:")
        assert js("source.value").endswith("\n    ")
        run_code("raise ValueError('console test')")
        assert "ValueError: console test" in rendered(".entry:last-child .error")
        js("clear.click()")
        wait_for(lambda: js("transcript.childElementCount === 0"))
        run_code("record['values'][0]")
        assert rendered(".result") == str(pid)
        cell = sheet.Range("N1")
        run_code("factor = 2\ndef scale(x, extra=1): return factor*x + extra")
        assert "=XLLPY.console.scale(x, extra)" in rendered(".entry:last-child .output")
        cell.Formula2 = "=XLLPY.console.scale(B1)"
        assert cell.Value2 == 85, cell.Text
        run_code("def scale(x, extra=1): return x + extra")
        wait_for(lambda: cell.Value2 == 43)
        run_code("def scale(x): return 3*x")
        wait_for(lambda: cell.Value2 == 126)
        run_code("del scale")
        wait_for(lambda: cell.Text in ("#NAME?", "#VALUE!"))
        run_code("def scale(x): return x")
        wait_for(lambda: cell.Value2 == 42)
        print("PASS: a worksheet cell calls a console function and tracks redefinition, signature changes and deletion", flush=True)
        marker = base / "finished-after-close"
        type_code(f"from pathlib import Path\nprint('started')\ntime.sleep(2)\nPath({str(marker)!r}).touch()")
        js("run.click()")
        wait_for(lambda: rendered(".entry:last-child .output").startswith("started"))
        print("PASS: ribbon console executes in Excel, streams output, expands nested objects, highlights, completes and retains variables", flush=True)


def verify_reopened_console(base):
    with page(base/"profile"/"ConsoleWebView2","https://python-console.example/") as (js, send):
        wait_for(lambda: js("!!document.querySelector('#source')"))
        assert js("transcript.childElementCount") == 0
        js("source.focus()")
        send("Input.insertText", text="'record' in globals()")
        js("run.click()")
        wait_for(lambda: js("transcript.querySelector('.result')?.textContent === 'False'"))
        wait_for(lambda: (base / "finished-after-close").exists())
        print("PASS: closing an active console lets its command finish and reopening starts a fresh session", flush=True)


def verify_editor(ide, script, read_console, action, pick_file):
    base=script.parent
    with page(base/"profile"/"WebView2","https://vscode.dev/") as (js, send):
        _exercise_editor(ide, base, read_console, action, pick_file, js, send)


def _exercise_editor(ide, base, read_console, action, pick_file, js, send):
    def shortcut(key):
        vk=ord(key.upper())
        for typ in ("keyDown","keyUp"):
            send("Input.dispatchKeyEvent",type=typ,key=key,code="Key"+key.upper(),modifiers=2,
                 windowsVirtualKeyCode=vk,nativeVirtualKeyCode=vk)
    try:
        wait_for(lambda:js("!!document.querySelector('.monaco-workbench')"),75)
        print("PASS: hosted vscode.dev workbench rendered inside Excel WebView2",flush=True)
        edited=base/"editor_saved.py"
        edited.write_text("print('before edit')\n",encoding="utf-8")
        send("Page.bringToFront")
        open_file="Array.from(document.querySelectorAll('button')).find(e=>e.textContent.trim()==='Open File...')"
        wait_for(lambda:js("Boolean("+open_file+")"))
        send("Runtime.evaluate",expression=open_file+".click()",userGesture=True)
        pick_file(edited,True)
        wait_for(lambda:js("document.title").find("editor_saved.py")>=0,25)
        wait_for(lambda:ide.child_window(title=str(edited),control_type="Text").exists(timeout=.2),10)
        shortcut("a")
        send("Input.insertText",text="print('VS_CODE_SAVE_AND_RUN')\n")
        shortcut("s")
        def saved():
            if "VS_CODE_SAVE_AND_RUN" in edited.read_text():
                return True
            for name in ("Save changes","Allow","Allow this time"):
                prompt=ide.child_window(title=name,control_type="Button")
                if prompt.exists(timeout=.1):
                    prompt.invoke()
            return False
        wait_for(saved,15)
        action("Run saved script")
        wait_for(lambda:"VS_CODE_SAVE_AND_RUN" in read_console(ide))
        print("PASS: open local file from VS Code, edit, save, automatically select and run in Excel CPython",flush=True)
    except Exception:
        print("Editor state:",js("document.body.innerText.slice(-1800)"),flush=True)
        raise

