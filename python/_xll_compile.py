"""Build a single project XLL from the shipping Excel-DNA host on Windows."""
import ast
import base64
import ctypes
from ctypes import wintypes
from hashlib import sha256
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import urllib.request
import uuid
import xml.etree.ElementTree as ET
import zipfile

SUPPORT = Path(__file__).resolve().parent


def inside(root, name):
    path = (root / name).resolve(strict=True)
    if Path(name).is_absolute() or not path.is_relative_to(root) or not path.is_file():
        raise ValueError(f"Expected a project file: {name}")
    return path


def command_names(root, entries):
    """Offer simple command candidates without importing the project in the designer."""
    names = []
    for entry in entries:
        module = list(Path(entry).with_suffix("").parts)
        if module[-1] == "__init__":
            module.pop()
        for fn in ast.parse(inside(root, entry).read_bytes(), entry).body:
            if not isinstance(fn, (ast.AsyncFunctionDef, ast.FunctionDef)) or fn.name.startswith("_"):
                continue
            args = fn.args
            if len(args.posonlyargs) + len(args.args) == len(args.defaults) and None not in args.kw_defaults:
                names.append(".".join([*module, fn.name]))
    return sorted(set(names))


def resource_bytes(template, name, kind):
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    def api(symbol, result, *args):
        fn = getattr(kernel, symbol)
        fn.argtypes, fn.restype = args, result
        return fn
    handle, string = wintypes.HANDLE, wintypes.LPCWSTR
    module = api("LoadLibraryExW", handle, string, handle, wintypes.DWORD)(str(template), None, 2)
    if not module:
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        entry = api("FindResourceW", handle, handle, string, string)(module, name, kind)
        if not entry:
            raise ValueError(f"The template is missing its {name} resource")
        size = api("SizeofResource", wintypes.DWORD, handle, handle)(module, entry)
        data = api("LoadResource", handle, handle, handle)(module, entry)
        pointer = api("LockResource", ctypes.c_void_p, handle)(data)
        if not pointer:
            raise ctypes.WinError(ctypes.get_last_error())
        return ctypes.string_at(pointer, size)
    finally:
        api("FreeLibrary", wintypes.BOOL, handle)(module)


def packed_dna(template, name):
    xml = ET.fromstring(resource_bytes(template, "__MAIN__", "DNA"))
    # The native Excel-DNA bootstrap scans for an unprefixed DnaLibrary tag.
    ET.register_namespace("", xml.tag[1:].split("}")[0])
    xml.set("Name", name)
    return ET.tostring(xml, encoding="utf-8", xml_declaration=True)


def optimised_template(workbench, output):
    pin = json.loads(resource_bytes(workbench, "RUNTIME_TEMPLATE", ctypes.cast(10, wintypes.LPCWSTR)))
    tag, digest = pin["tag"], pin["sha256"]
    if not re.fullmatch(r"v[0-9]+\.[0-9]+\.[0-9]+", tag) or not re.fullmatch(r"[a-f0-9]{64}", digest):
        raise ValueError("Invalid runtime template version or SHA-256 in the workbench")
    asset = "PythonExcelRuntime.xll"
    root = Path(os.environ["XLL_PYTHON_HOME"]) if "XLL_PYTHON_HOME" in os.environ else Path(os.environ["LOCALAPPDATA"]) / "XllPython"
    cache = root / "templates" / tag / digest / asset
    sibling = workbench.with_name(asset)
    if output in (cache.resolve(), sibling.resolve()):
        raise ValueError("Choose an output file separate from the runtime template")
    for path in (cache, sibling):
        if path.is_file():
            data = path.read_bytes()
            if sha256(data).hexdigest() == digest:
                return data
    url = f"https://github.com/sancarn/xll-python/releases/download/{tag}/{asset}"
    try:
        with urllib.request.urlopen(url, timeout=90) as response:
            data = response.read()
    except OSError as error:
        raise ValueError(f"Could not download the runtime template for {tag}. "
                         f"For offline compilation, place the matching {asset} beside {workbench.name}. {error}") from error
    if sha256(data).hexdigest() != digest:
        raise ValueError(f"The runtime template for {tag} failed SHA-256 verification")
    cache.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(dir=cache.parent) as temp:
        candidate = Path(temp) / asset
        candidate.write_bytes(data)
        try:
            os.replace(candidate, cache)
        except PermissionError:
            # Another compiler may already have published this copy and opened it for reading.
            if not cache.is_file() or sha256(cache.read_bytes()).hexdigest() != digest:
                raise
    return data


def resources(target, items):
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    begin = kernel.BeginUpdateResourceW
    begin.argtypes = [wintypes.LPCWSTR, wintypes.BOOL]
    begin.restype = wintypes.HANDLE
    update = kernel.UpdateResourceW
    update.argtypes = [wintypes.HANDLE, ctypes.c_void_p, wintypes.LPCWSTR, wintypes.WORD,
                       ctypes.c_void_p, wintypes.DWORD]
    end = kernel.EndUpdateResourceW
    end.argtypes = [wintypes.HANDLE, wintypes.BOOL]
    handle = begin(str(target), False)
    if not handle:
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        for name, data in items.items():
            kind = ctypes.cast(ctypes.c_wchar_p("DNA"), ctypes.c_void_p) if name == "__MAIN__" else 10
            if not update(handle, kind, name, 0, data, len(data) if data is not None else 0):
                raise ctypes.WinError(ctypes.get_last_error())
    except BaseException:
        end(handle, True)
        raise
    if not end(handle, False):
        raise ctypes.WinError(ctypes.get_last_error())


def ribbon(project, root, available):
    ns = "http://schemas.microsoft.com/office/2009/07/customui"
    prefix = "p" + uuid.UUID(project["id"]).hex
    xml = ET.Element("customUI", xmlns=ns, loadImage="LoadImage")
    tabs = ET.SubElement(ET.SubElement(xml, "ribbon"), "tabs")
    actions, images = {}, {}
    for t, tab in enumerate(project["tabs"]):
        tid = f"{prefix}t{t}"
        tnode = ET.SubElement(tabs, "tab", id=tid, label=tab["label"])
        for g, group in enumerate(tab["groups"]):
            gid = f"{tid}g{g}"
            gnode = ET.SubElement(tnode, "group", id=gid, label=group["label"])
            for b, button in enumerate(group["buttons"]):
                bid = f"{gid}b{b}"
                target = button["command"]
                if target not in available:
                    raise ValueError(f"{button['label']}: {target} must be callable without arguments")
                opts = dict(id=bid, label=button["label"], onAction="RunPublished",
                            size="large" if button.get("large") else "normal")
                if button.get("image"):
                    data = inside(root, button["image"]).read_bytes()
                    if not data.startswith(b"\x89PNG\r\n\x1a\n"):
                        raise ValueError(f"Expected a PNG image for {button['label']}")
                    image = "i" + sha256(data).hexdigest()
                    opts["image"] = image
                    images[image] = base64.b64encode(data).decode("ascii")
                elif button.get("imageMso"):
                    opts["imageMso"] = button["imageMso"]
                ET.SubElement(gnode, "button", **opts)
                actions[bid] = target
    return dict(ribbon=ET.tostring(xml, encoding="unicode") if project["tabs"] else "",
                actions=actions, images=images)


def publish(project_file, template, output):
    project_file, template, output = map(lambda p: Path(p).resolve(), (project_file, template, output))
    root = project_file.parent
    project = json.loads(project_file.read_text(encoding="utf-8"))
    if not isinstance(project.get("optimise", False), bool):
        raise ValueError("Optimise must be true or false")
    if output == template or output.suffix.lower() != ".xll":
        raise ValueError("Choose a separate .xll output file")
    if not project["entries"] or not project["name"].strip():
        raise ValueError("Enter an add-in name and select at least one Python entry file")
    with tempfile.TemporaryDirectory(dir=output.parent, prefix=".xll-build-") as temp:
        stage = Path(temp) / "project"
        for name in dict.fromkeys(project["entries"] + project["files"]):
            path = inside(root, name)
            dest = stage / "src" / path.relative_to(root)
            dest.parent.mkdir(exist_ok=True, parents=True)
            shutil.copyfile(path, dest)
        (stage / "project.json").write_text(json.dumps(dict(id=project["id"], entries=project["entries"])), encoding="utf-8")
        requirements = []
        wheels = []
        for spec in project["packages"]:
            if spec.lower().endswith(".whl"):
                wheel = root / spec
                # Search beside the wheel the author selected, including a symlinked wheel.
                if not wheel.is_file():
                    raise ValueError(f"Expected a wheel file: {spec}")
                requirements.append(str(wheel))
                wheels.extend(["--find-links", str(wheel.parent)])
            elif re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._-]*(==[A-Za-z0-9][A-Za-z0-9.!+_-]*)?", spec):
                requirements.append(spec)
            else:
                raise ValueError(f"Expected a package name, exact version or local wheel: {spec}")
        if requirements:
            subprocess.run([sys.executable, "-I", "-m", "pip", "--disable-pip-version-check", "install",
                            "--only-binary=:all:", "--ignore-installed", "--no-compile", "--target",
                            str(stage / "packages"), *wheels, *requirements], check=True)
        metadata = Path(temp) / "exports.json"
        subprocess.run([sys.executable, "-I", "-S", "-B", str(SUPPORT / "_xll_worker.py"),
                        "inspect", str(stage), str(metadata)], check=True)
        catalog = json.loads(metadata.read_text(encoding="utf-8"))
        ui = ribbon(project, root, catalog["commands"])
        if not catalog["functions"] and not ui["actions"]:
            raise ValueError("Select a ribbon command or export at least one worksheet function")
        archive = Path(temp) / "project.zip"
        with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as z:
            for path in sorted(stage.rglob("*")):
                if path.is_file():
                    z.write(path, path.relative_to(stage))
        data = archive.read_bytes()
        info = dict(name=project["name"], id=project["id"], hash=sha256(data).hexdigest(),
                    functions=catalog["functions"], **ui)
        built = Path(temp) / "output.xll"
        updates = {"PYTHON_PROJECT": data, "PYTHON_MANIFEST": json.dumps(info).encode()}
        if project.get("optimise", False):
            built.write_bytes(optimised_template(template, output))
        else:
            shutil.copyfile(template, built)
            updates["EXCEL_IDE"] = None
        updates["__MAIN__"] = packed_dna(built, project["name"])
        resources(built, updates)
        os.replace(built, output)


if __name__ == "__main__":
    publish(*sys.argv[1:])
