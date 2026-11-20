"""Exercise the real vscode.dev editor using only the disposable WebView2 profile."""
import json
from pathlib import Path
import urllib.request

import websocket
from excel_smoke import wait_for


def verify_editor(ide, script, read_console, action, pick_file):
    base=script.parent
    def port_file():
        return next((base/"profile"/"WebView2").rglob("DevToolsActivePort"),None)
    port=int(wait_for(port_file).read_text().splitlines()[0])
    def target():
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/list",timeout=3) as r:
            return next((t for t in json.load(r) if t["type"]=="page" and t["url"].startswith("https://vscode.dev/")),None)
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
        return send("Runtime.evaluate",expression=expression,returnByValue=True)["result"].get("value")
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
    finally:
        ws.close()

