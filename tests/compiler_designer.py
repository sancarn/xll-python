"""Drive the real Tk designer with deterministic file-dialog selections in a disposable project."""
import json
from pathlib import Path
import sys
import time
import tkinter as tk
from tkinter import ttk

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "python"))
import _xll_designer as ui


def main(project, template, output):
    window = tk.Tk()
    app = ui.Designer(window, template)
    ui.filedialog.askopenfilename = lambda **kw: str(project)
    app.open()
    # Use the same action as "Include selected" in the Packages tab.
    spec = "xll_fixture==1"
    index = app.installed.get(0, "end").index(spec)
    app.installed.selection_set(index)
    app.packages.delete(0, "end")
    def widgets(parent):
        for child in parent.winfo_children():
            yield child
            yield from widgets(child)
    def button(parent, label):
        return next(w for w in widgets(parent) if isinstance(w, ttk.Button) and w["text"] == label)
    button(window, "Include selected").invoke()
    app.tabs.clear()
    app.render()
    names = iter(["First", "Reports"])
    ui.simpledialog.askstring = lambda *a, **kw: next(names)
    button(window, "Add tab").invoke()
    button(window, "Add group").invoke()
    button(window, "Add button").invoke()
    dialog = next(w for w in window.winfo_children() if isinstance(w, tk.Toplevel))
    label, command, office = [w for w in widgets(dialog) if isinstance(w, ttk.Entry)]
    label.insert(0, "Refresh First")
    assert "reports.refresh" in command["values"]
    command.set("reports.refresh")
    ui.filedialog.askopenfilename = lambda **kw: str(project.parent / "icon.png")
    button(dialog, "Import PNG…").invoke()
    button(dialog, "Save button").invoke()
    button(window, "Save project").invoke()
    saved = json.loads(project.read_text(encoding="utf-8"))
    assert saved["packages"] == [spec] and saved["tabs"][0]["groups"][0]["buttons"][0]["image"].startswith("ribbon-images/")
    ui.filedialog.askopenfilename = lambda **kw: str(project)
    app.open()
    assert app.state() == saved
    ui.filedialog.asksaveasfilename = lambda **kw: str(output)
    problems = []
    ui.messagebox.showerror = lambda *args, **kw: problems.append(args)
    button(window, "Compile XLL…").invoke()
    end = time.monotonic() + 90
    while app.process and time.monotonic() < end:
        window.update()
        time.sleep(.02)
    assert output.is_file() and not problems, problems
    window.destroy()


if __name__ == "__main__":
    main(*map(Path, sys.argv[1:]))
