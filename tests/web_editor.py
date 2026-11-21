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
        run_code("from xll_python import Application\ncell_ref = Application.ActiveSheet.Range(\"T3\")")
        run_code("cell_ref.Value2 = 73")
        assert sheet.Range("T3").Value2 == 73
        run_code("import contextvars\nsaved = contextvars.copy_context()\ndef guarded(): return saved.run(lambda: Application.ActiveCell.Value2)")
        sheet.Range("T4").Formula2 = "=XLLPY.console.guarded()"
        assert sheet.Range("T4").Text == "#VALUE!"
        run_code("nested = 1\nfor _ in range(150): nested = [nested]\ntry:\n    cell_ref.Value2 = nested\nexcept RuntimeError:\n    cell_ref.Value2 = 74")
        assert sheet.Range("T3").Value2 == 74
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


def verify_packages(base, install):
    with page(base/"profile"/"PackagesWebView2", "https://python-packages.example/") as (js, send):
        wait_for(lambda: js("document.querySelector('[aria-label=\"Add openpyxl\"]') !== null"))
        js("document.querySelector('[aria-label=\"Add openpyxl\"]').click();"
           "document.querySelector('[aria-label=\"Add xlsxwriter\"]').click()")
        wait_for(lambda: js("document.querySelector('[data-name=openpyxl] .running') && "
                            "document.querySelector('[data-name=xlsxwriter] .actions').innerText === 'Queued'"))
        js("document.querySelector('#advanced-tab').click(); document.querySelector('#command').focus()")
        send("Input.insertText", text="pip " + install)
        js("document.querySelector('#prompt').requestSubmit()")
        wait_for(lambda: js("document.querySelectorAll('.job').length === 3"))


def verify_package_log(base):
    with page(base/"profile"/"PackagesWebView2", "https://python-packages.example/") as (js, send):
        wait_for(lambda: js("document.querySelectorAll('#results .done').length === 2"))
        js("document.querySelector('#installed-tab').click()")
        wait_for(lambda: js("document.querySelector('#distributions').innerText.toLowerCase().replaceAll('_','-').includes('xll-probe')"))
        js("document.querySelector('#advanced-tab').click()")
        assert js("Array.from(document.querySelectorAll('.job code'), e => e.innerText).slice(0, 2)") == [
            "pip install openpyxl", "pip install xlsxwriter"]
        assert js("document.querySelectorAll('.job .done').length") == 3
        assert js("document.querySelector('.job:last-child pre').innerText.includes('Successfully installed')")
        js("document.querySelector('#browse-tab').click(); document.querySelector('[aria-label=\"Add numpy\"]').click()")
        wait_for(lambda: js("!!document.querySelector('[aria-label=\"Retry numpy\"]')"))
        js("document.querySelector('[aria-label=\"View log for numpy\"]').click()")
        assert js("document.activeElement.matches('.job:last-child') && document.activeElement.innerText.includes('No matching distribution')")
        from excel_ui import fixture_wheel
        fixture_wheel(base, "numpy")
        js("document.querySelector('#browse-tab').click(); document.querySelector('[aria-label=\"Retry numpy\"]').click()")
        wait_for(lambda: js("!!document.querySelector('[data-name=numpy] .done')"))
        js("document.querySelector('#search').focus()")
        send("Input.insertText", text="beautifulsoup4")
        wait_for(lambda: js("document.querySelector('#results .row')?.dataset.name.toLowerCase() === 'beautifulsoup4'"), 90)
        wait_for(lambda: js("!!document.querySelector('#results .row p')?.innerText"))
        print("PASS: queued Browse and Advanced installs survived reopen, inventory and logs agree, failed install links to its log, live PyPI search returns metadata", flush=True)



def close_during_attach(base, close):
    import socket

    # Accept the real debug connection, but never send configurationDone.
    with socket.create_server(("127.0.0.1", 0)) as adapter:
        adapter.settimeout(15)
        with page(base / "profile" / "IdeWebView2", "http://xll-") as (js, _):
            def initialized():
                js("[...document.querySelectorAll('button')].find(b => b.innerText.includes('Yes, I trust'))?.click()")
                return "Python 3.13" in js("document.body.innerText").replace("\xa0", " ")
            wait_for(initialized)
            request = dict(id=-1, operation="connect", args=dict(port=adapter.getsockname()[1], token="test"))
            js(f"window.chrome.webview.postMessage({json.dumps(request)})")
            client, _ = adapter.accept()
            with client:
                close()
                client.settimeout(15)
                try:
                    while client.recv(4096):
                        pass
                except ConnectionResetError:
                    pass


def verify_editor(base, pid, close):
    """Exercise the shipped workbench and a synchronous cell in this Excel instance."""
    import queue
    import threading
    import pythoncom
    from excel_smoke import connect

    script = base / "scripts" / "ide_runner.py"
    helper = base / "packages" / "ide_helper.py"
    helper.write_text("def twice(value):\n    factor = 2\n    return value * factor\n")
    source = ("from xll_python import function\nfrom ide_helper import twice\n"
              "@function(name='PY.IDE_DOUBLE')\ndef calc(value: float): return twice(value)\n")
    script.write_text(source)
    (base / "scripts" / "ide_streams.py").write_text(
        "from xll_python import function\nfrom ide_helper import twice\n"
        "from pathlib import Path\nimport asyncio, time\n"
        "@function(name='PY.IDE_SYNC')\ndef sync():\n"
        "    value = 1\n    flag = Path(__file__).with_name('ide-sync')\n"
        "    while True:\n        if flag.exists():\n            flag.unlink()\n            value = twice(17)\n"
        "        yield value\n        time.sleep(.15)\n"
        "@function(name='PY.IDE_ASYNC')\nasync def asyn():\n"
        "    value = 1\n    flag = Path(__file__).with_name('ide-async')\n"
        "    while True:\n        if flag.exists():\n            flag.unlink()\n            value = twice(18)\n"
        "        yield value\n        await asyncio.sleep(.15)\n"
    )
    result = queue.Queue()

    def calculate(value):
        pythoncom.CoInitialize()
        try:
            app = connect(pid, base / "book.csv")
            cell = app.ActiveSheet.Range("Q9")
            cell.Formula2 = f"=PY.IDE_DOUBLE({value})"
            result.put(cell.Value2)
        except BaseException as error:
            result.put(error)
        finally:
            pythoncom.CoUninitialize()

    with page(base / "profile" / "IdeWebView2", "http://xll-") as (js, send):
        def body():
            return js("document.body.innerText").replace("\xa0", " ")
        def key(name, vk, modifiers=0):
            for event in ("keyDown", "keyUp"):
                send("Input.dispatchKeyEvent", type=event, key=name, windowsVirtualKeyCode=vk,
                     nativeVirtualKeyCode=vk, modifiers=modifiers)
        def select(text, command=False):
            key("F1", 112) if command else key("p", 80, 2)
            wait_for(lambda: js("!!document.querySelector('.quick-input-widget input:focus')"), 8)
            send("Input.insertText", text=text)
            wait_for(lambda: js("document.querySelector('.quick-input-list')?.innerText ?? ''").find(text) >= 0)
            key("Enter", 13)
        def initialized():
            js("""[...document.querySelectorAll('button')].find(b => b.innerText.includes('Yes, I trust'))?.click()""")
            return f"Excel {pid} | Python 3.13" in body()

        wait_for(initialized)
        send("Page.bringToFront")
        select("Excel: Refresh Worksheet Functions", True)
        wait_for(lambda: "Refreshed Excel worksheet functions." in body())
        app = connect(pid, base / "book.csv")
        sheet = app.ActiveSheet
        sheet.Range("R8").Formula2 = "=PY.IDE_SYNC()"
        sheet.Range("R9").Formula2 = "=PY.IDE_ASYNC()"
        wait_for(lambda: sheet.Range("R8:R9").Value2 == ((1.,), (1.,)))
        select(script.name)
        wait_for(lambda: "Ln " in body() and script.name in body())
        key("End", 35, 2)
        send("Input.insertText", text="import xll_python\nxll_python.")
        key(" ", 32, 2)
        wait_for(lambda: {"Reference", "ExcelError"} <= set(body().splitlines()))
        key("Escape", 27)
        key("a", 65, 2)
        send("Input.insertText", text=source + "if __name__ == '__main__':\n"
             "    from xll_python import Application\n"
             "    Application.ActiveSheet.Range('U3').Value2 = 'IDE'\n"
             "print('IDE_RUN', __import__('os').getpid())\n")
        key("F5", 116, 2)
        wait_for(lambda: f"IDE_RUN {pid}" in body())
        assert "print('IDE_RUN'" in script.read_text()
        assert sheet.Range("U3").Value2 == "IDE"
        print("PASS: Pyright completed xll_python; Run saved the active editor and executed inside Excel.", flush=True)

        select(script.name)
        key("End", 35, 2)
        key("ArrowUp", 38)
        key("F9", 120)
        wait_for(lambda: js("!!document.querySelector('.cgmr.codicon-debug-breakpoint')"))
        key("F5", 116)
        wait_for(lambda: "PAUSED ON BREAKPOINT" in body())
        key("F5", 116)
        wait_for(lambda: js("document.querySelector('#outputView')?.textContent ?? ''").count(f"IDE_RUN\xa0{pid}") >= 2)
        wait_for(lambda: "PAUSED ON BREAKPOINT" not in body())
        print("PASS: F5 stopped in the active script and continued in the same Excel process.", flush=True)

        select(helper.name)
        wait_for(lambda: "return value * factor" in body())
        key("End", 35, 2)
        key("ArrowUp", 38)
        key("F9", 120)
        wait_for(lambda: js("!!document.querySelector('.cgmr.codicon-debug-breakpoint')"))
        select("Excel: Debug Worksheet Functions", True)
        wait_for(lambda: "RUNNING" in body())
        for signal, value, cell in (("ide-sync", 17, "R8"), ("ide-async", 18, "R9")):
            (base / "scripts" / signal).touch()
            wait_for(lambda: "PAUSED ON BREAKPOINT" in body() and "factor:" in body())
            key("F5", 116)
            wait_for(lambda: sheet.Range(cell).Value2 == value * 2)
        sheet.Range("R8:R9").ClearContents()
        print("PASS: attached to pre-existing sync and async generator workers.", flush=True)
        for value in (21, 22):
            threading.Thread(target=calculate, args=(value,), daemon=True).start()
            wait_for(lambda: "PAUSED ON BREAKPOINT" in body())
            wait_for(lambda: "factor:" in body() and str(value) + ".0" in body())
            if value == 21:
                key("F5", 116)
                assert result.get(timeout=12) == 42
                key("F5", 116, 8)
                wait_for(lambda: "RUNNING" not in body())
                select("Excel: Debug Worksheet Functions", True)
                wait_for(lambda: "RUNNING" in body())
            else:
                close()
                assert result.get(timeout=12) == 44
    print("PASS: inspected a helper in a second root, continued, stopped, reattached and closed while paused.", flush=True)
