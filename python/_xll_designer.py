"""Project and ribbon editor launched by the workbench's Compile to XLL command."""
from hashlib import sha256
from importlib import metadata
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
import uuid

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from _xll_compile import command_names


def row(parent):
    frame = ttk.Frame(parent)
    frame.pack(fill="x", pady=4)
    return frame


def button(parent, label, action):
    ttk.Button(parent, text=label, command=action).pack(side="left", padx=(0, 5))


def values(parent, height=7):
    box = tk.Listbox(parent, selectmode="extended", exportselection=False, height=height)
    box.pack(fill="both", expand=True, pady=3)
    return box


def selected(box):
    return [box.get(i) for i in box.curselection()]


def add(box, items):
    existing = set(box.get(0, "end"))
    for item in items:
        if item not in existing:
            box.insert("end", item)
            existing.add(item)


class Designer:
    def __init__(self, app, template):
        self.app, self.template = app, template
        self.file = self.process = None
        self.app.title("Compile to XLL")
        self.app.geometry("780x650")
        self.app.minsize(620, 530)
        self.app.report_callback_exception = lambda typ, error, tb: messagebox.showerror("Compile to XLL", str(error))
        self.app.protocol("WM_DELETE_WINDOW", self.close)
        main = ttk.Frame(app, padding=14)
        main.pack(fill="both", expand=True)
        toolbar = row(main)
        button(toolbar, "New project", self.new)
        button(toolbar, "Open project", self.open)
        button(toolbar, "Save project", self.save)
        button(toolbar, "Compile XLL…", self.compile)
        title = row(main)
        ttk.Label(title, text="Add-in name").pack(side="left", padx=(0, 8))
        self.name = tk.StringVar()
        ttk.Entry(title, textvariable=self.name).pack(side="left", fill="x", expand=True)
        tabs = ttk.Notebook(main)
        tabs.pack(fill="both", expand=True, pady=6)
        sources, ribbon, packages = [ttk.Frame(tabs, padding=10) for _ in range(3)]
        for frame, label in [(sources,"Project files"), (ribbon,"Ribbon"), (packages,"Packages")]:
            tabs.add(frame, text=label)
        ttk.Label(sources, text="Entry modules").pack(anchor="w")
        self.entries = values(sources)
        bar = row(sources)
        button(bar, "Add Python files…", lambda: self.choose_files(self.entries, True))
        button(bar, "Remove", lambda: self.remove(self.entries))
        ttk.Label(sources, text="Helper modules and data files").pack(anchor="w", pady=(8, 0))
        self.files = values(sources)
        bar = row(sources)
        button(bar, "Add files…", lambda: self.choose_files(self.files, False))
        button(bar, "Remove", lambda: self.remove(self.files))

        self.tree = ttk.Treeview(ribbon, show="tree headings", columns=("command",), selectmode="browse")
        self.tree.heading("#0", text="Tab / Group / Button")
        self.tree.heading("command", text="Python function")
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Double-1>", lambda _: self.edit())
        bar = row(ribbon)
        for label, kind in [("Add tab","tabs"), ("Add group","groups"), ("Add button","buttons")]:
            button(bar, label, lambda k=kind: self.append(k))
        button(bar, "Edit", self.edit)
        button(bar, "Remove", self.remove_node)

        ttk.Label(packages, text="Installed in the workbench").pack(anchor="w")
        self.installed = values(packages)
        add(self.installed, sorted({d.metadata["Name"] + "==" + d.version for d in metadata.distributions()
                                    if d.metadata["Name"]}, key=str.lower))
        bar = row(packages)
        button(bar, "Include selected", lambda: add(self.packages, selected(self.installed)))
        button(bar, "Add by name…", self.add_package)
        button(bar, "Add local wheel…", self.wheel)
        ttk.Label(packages, text="Packages to bundle (dependencies are included automatically)").pack(anchor="w", pady=(8, 0))
        self.packages = values(packages)
        button(row(packages), "Remove", lambda: self.remove(self.packages))
        self.status = tk.StringVar(value="Create or open a project to choose files.")
        ttk.Label(main, textvariable=self.status, wraplength=720).pack(anchor="w")
        self.reset()

    def reset(self, project=None):
        p = project or dict(id=str(uuid.uuid4()), name="", entries=[], files=[], packages=[], tabs=[])
        self.id = p["id"]
        self.name.set(p["name"])
        self.tabs = p["tabs"]
        for box, key in [(self.entries,"entries"), (self.files,"files"), (self.packages,"packages")]:
            box.delete(0, "end")
            add(box, p[key])
        self.render()
        self.saved = self.state()

    def state(self):
        return dict(id=self.id, name=self.name.get(), entries=list(self.entries.get(0, "end")),
                    files=list(self.files.get(0, "end")), packages=list(self.packages.get(0, "end")),
                    tabs=json.loads(json.dumps(self.tabs)))

    def leave(self):
        if self.process:
            messagebox.showinfo("Compile to XLL", "Wait for the current build to finish.")
            return False
        if self.state() != self.saved:
            answer = messagebox.askyesnocancel("Compile to XLL", "Save changes to this project?")
            if answer is None or (answer and not self.save()):
                return False
        return True

    def close(self):
        if self.leave():
            self.app.destroy()

    def new(self):
        if not self.leave():
            return
        path = filedialog.asksaveasfilename(title="Create project beside your Python files",
            defaultextension=".xllproject.json", filetypes=[("XLL project","*.xllproject.json")])
        if path:
            self.file = Path(path)
            self.reset()
            self.name.set(self.file.name.split(".")[0])
            self.save()

    def open(self):
        if not self.leave():
            return
        path = filedialog.askopenfilename(filetypes=[("XLL project","*.json")])
        if path:
            p = json.loads(Path(path).read_text(encoding="utf-8"))
            self.reset(p)
            self.file = Path(path)
            self.status.set(str(self.file))

    def save(self):
        if not self.file:
            path = filedialog.asksaveasfilename(title="Save project beside your Python files",
                defaultextension=".xllproject.json", filetypes=[("XLL project","*.xllproject.json")])
            if not path:
                return False
            self.file = Path(path)
        state = self.state()
        temp = self.file.with_name(self.file.name + "." + uuid.uuid4().hex + ".tmp")
        try:
            temp.write_text(json.dumps(state, indent=2) + "\n", encoding="utf-8")
            os.replace(temp, self.file)
        finally:
            temp.unlink(missing_ok=True)
        self.saved = state
        self.status.set(str(self.file))
        return True

    def root(self):
        if self.file is None and not self.save():
            return None
        return self.file.parent

    def choose_files(self, box, python):
        root = self.root()
        if not root:
            return
        paths = filedialog.askopenfilenames(initialdir=root, filetypes=[("Python","*.py")] if python else [("All files","*.*")])
        for path in map(Path, paths):
            if not path.is_relative_to(root):
                raise ValueError("Place project files beneath the project directory: " + str(root))
            add(box, [path.relative_to(root).as_posix()])

    def copy_asset(self, path, folder):
        root = self.file.parent
        src = Path(path)
        if folder == "wheels" and src.is_relative_to(root):
            return src.relative_to(root).as_posix()
        dest = root / folder / (sha256(src.read_bytes()).hexdigest()[:16] + ".png" if folder == "ribbon-images" else src.name)
        dest.parent.mkdir(exist_ok=True)
        if src.resolve() != dest.resolve():
            if dest.exists() and dest.read_bytes() != src.read_bytes():
                raise ValueError("A different project file already exists at " + str(dest))
            if not dest.exists():
                shutil.copyfile(src, dest)
        return dest.relative_to(root).as_posix()

    def wheel(self):
        if not self.root():
            return
        for path in filedialog.askopenfilenames(filetypes=[("Python wheels","*.whl")]):
            add(self.packages, [self.copy_asset(path, "wheels")])

    def add_package(self):
        spec = simpledialog.askstring("Include package", "Package name or name==version")
        if spec and spec.strip():
            add(self.packages, [spec.strip()])

    @staticmethod
    def remove(box):
        for index in reversed(box.curselection()):
            box.delete(index)

    def render(self, focus=None):
        self.tree.delete(*self.tree.get_children())
        self.nodes = {}
        def insert(parent, items, kind):
            for n, item in enumerate(items):
                key = f"{parent}_{n}"
                self.nodes[key] = (item, items, kind)
                self.tree.insert(parent, "end", iid=key, text=item["label"], open=True,
                                 values=(item.get("command", ""),))
                if item is focus:
                    self.tree.selection_set(key)
                if kind != "buttons":
                    child = "groups" if kind == "tabs" else "buttons"
                    insert(key, item[child], child)
        insert("", self.tabs, "tabs")

    def node(self):
        keys = self.tree.selection()
        return self.nodes[keys[0]] if keys else (None, None, None)

    def append(self, kind):
        item, _, parent_kind = self.node()
        if kind == "tabs":
            items = self.tabs
        else:
            expected = "tabs" if kind == "groups" else "groups"
            if parent_kind != expected:
                raise ValueError("Select a tab first." if expected == "tabs" else "Select a group first.")
            items = item[kind]
        if kind == "buttons":
            self.button_dialog({}, items)
        else:
            label = simpledialog.askstring("Add " + kind[:-1], "Display name")
            if label and label.strip():
                new = dict(label=label.strip())
                new["groups" if kind == "tabs" else "buttons"] = []
                items.append(new)
                self.render(new)

    def edit(self):
        item, items, kind = self.node()
        if item is None:
            return
        if kind == "buttons":
            self.button_dialog(item, items)
        else:
            label = simpledialog.askstring("Rename", "Display name", initialvalue=item["label"])
            if label and label.strip():
                item["label"] = label.strip()
                self.render(item)

    def remove_node(self):
        item, items, _ = self.node()
        if item is not None:
            items.remove(item)
            self.render()

    def button_dialog(self, item, items):
        root = self.root()
        if not root:
            return
        candidates = command_names(root, self.entries.get(0, "end"))
        dialog = tk.Toplevel(self.app)
        dialog.title("Ribbon button")
        dialog.transient(self.app)
        dialog.grab_set()
        form = ttk.Frame(dialog, padding=12)
        form.pack(fill="both", expand=True)
        label = tk.StringVar(value=item.get("label", ""))
        command = tk.StringVar(value=item.get("command", ""))
        office = tk.StringVar(value=item.get("imageMso", "MacroPlay"))
        image = tk.StringVar(value=item.get("image", ""))
        large = tk.BooleanVar(value=item.get("large", True))
        for title, var, options in [("Button label",label,None), ("Python function",command,candidates),
                                    ("Office icon",office,["MacroPlay","RefreshAll","FileOpen","FileSave","TableInsert"])]:
            ttk.Label(form, text=title).pack(anchor="w")
            field = ttk.Entry(form, textvariable=var, width=45) if options is None else ttk.Combobox(form, textvariable=var, values=options, width=42)
            field.pack(fill="x", pady=(0, 6))
        preview = ttk.Label(form)
        def show(*_):
            preview.photo = None
            if image.get():
                photo = tk.PhotoImage(file=root / image.get())
                px = 32 if large.get() else 16
                factor = max(1, (max(photo.width(), photo.height()) + px - 1) // px)
                preview.photo = photo.subsample(factor)
            preview.configure(text=label.get(), image=preview.photo or "", compound="left")
        def choose():
            path = filedialog.askopenfilename(parent=dialog, filetypes=[("PNG image","*.png")])
            if path:
                image.set(self.copy_asset(path, "ribbon-images"))
        bar = row(form)
        button(bar, "Import PNG…", choose)
        button(bar, "Use Office icon", lambda: image.set(""))
        ttk.Checkbutton(form, text="Large button", variable=large).pack(anchor="w")
        preview.pack(pady=8)
        for var in (image, large, label):
            var.trace_add("write", show)
        show()
        def apply():
            if not label.get().strip() or not command.get().strip():
                messagebox.showerror("Ribbon button", "Enter a label and choose a Python function.", parent=dialog)
                return
            if not item:
                items.append(item)
            item.update(label=label.get().strip(), command=command.get().strip(),
                        image=image.get(), imageMso=office.get().strip(), large=large.get())
            self.render(item)
            dialog.destroy()
        button(row(form), "Save button", apply)

    def compile(self):
        if self.process or not self.save():
            return
        path = filedialog.asksaveasfilename(title="Save compiled add-in", defaultextension=".xll",
            initialdir=self.file.parent, initialfile=self.name.get()+".xll", filetypes=[("Excel add-in","*.xll")])
        if not path:
            return
        self.status.set("Building " + path)
        # Send verbose pip diagnostics to a file so a full pipe can never stall the UI.
        import tempfile
        log = tempfile.TemporaryFile()
        try:
            self.process = subprocess.Popen([Path(sys.executable).with_name("python.exe"), "-I",
                HERE / "_xll_compile.py", self.file, self.template, path],
                stdout=log, stderr=subprocess.STDOUT, creationflags=subprocess.CREATE_NO_WINDOW)
        except BaseException:
            log.close()
            raise
        def check():
            if self.process.poll() is None:
                self.app.after(150, check)
                return
            code = self.process.returncode
            self.process = None
            log.seek(0)
            detail = log.read().decode(errors="replace")
            log.close()
            self.status.set(("Build failed" if code else "Built " + path))
            if code:
                messagebox.showerror("Compile failed", detail[-4000:])
        self.app.after(150, check)


if __name__ == "__main__":
    window = tk.Tk()
    Designer(window, sys.argv[1])
    window.mainloop()
