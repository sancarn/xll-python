"""Filesystem operations for the local Compile to XLL WebView2 designer."""
import base64
from hashlib import sha256
from importlib import metadata
import json
import os
from pathlib import Path
import stat
import subprocess
import sys
import tempfile
import uuid

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from _xll_compile import command_names, inside


def source_files(root, manifest):
    paths = []
    for folder, dirs, files in os.walk(root):
        here = Path(folder)
        dirs[:] = sorted(d for d in dirs if not d.startswith(".") and
                         d not in {"__pycache__", "node_modules"} and
                         not (here / d / "pyvenv.cfg").is_file() and
                         not (here / d).lstat().st_file_attributes & stat.FILE_ATTRIBUTE_REPARSE_POINT)
        for name in files:
            path = here / name
            if (name.startswith(".") or name.lower().endswith((".xll", ".pyc", ".xllproject.json"))
                    or path == manifest or path.is_symlink() or not path.is_file()):
                continue
            paths.append(path.relative_to(root).as_posix())
    return sorted(paths, key=str.casefold)


def read_project(path):
    project = json.loads(path.read_text(encoding="utf-8-sig"))
    uuid.UUID(project["id"])
    if not isinstance(project["name"], str):
        raise ValueError("Expected an add-in name in the project file.")
    for key in ("entries", "files", "packages", "tabs"):
        if not isinstance(project[key], list):
            raise ValueError(f"Expected a list of {key} in the project file.")
    for key in ("entries", "files"):
        project[key] = [Path(name).as_posix() for name in project[key]]
    if not all(isinstance(spec, str) for spec in project["packages"]):
        raise ValueError("Expected package names or wheel paths.")

    def check_nodes(items, level):
        if not isinstance(items, list):
            raise ValueError("Expected a list of ribbon items.")
        for item in items:
            if not isinstance(item, dict) or not isinstance(item.get("label"), str):
                raise ValueError("Each ribbon item needs a display label.")
            if level == "buttons":
                if not isinstance(item.get("command"), str):
                    raise ValueError("Each ribbon button needs a Python function.")
            else:
                child = "groups" if level == "tabs" else "buttons"
                check_nodes(item[child], child)

    check_nodes(project["tabs"], "tabs")
    return project


def save_project(path, project):
    with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=path.parent, delete=False) as f:
        temp = Path(f.name)
    try:
        temp.write_text(json.dumps(project, indent=2) + "\n", encoding="utf-8")
        os.replace(temp, path)
    finally:
        temp.unlink(missing_ok=True)


def handle(req):
    action = req["action"]
    root = Path(req["root"]).resolve() if req.get("root") else None
    manifest = Path(req["manifest"]) if req.get("manifest") else None
    if action in ("new", "open", "refresh"):
        files = source_files(root, manifest.resolve() if manifest else None)
        if action == "refresh":
            return files
        project = read_project(manifest) if action == "open" else dict(
            id=str(uuid.uuid4()), name=root.name, entries=[], files=[p for p in files if p.lower().endswith(".py")],
            packages=[], tabs=[])
        return dict(project=project, paths=files, root=str(root), manifest=str(manifest or ""))
    if action == "save":
        if manifest.parent.resolve() != root:
            raise ValueError(f"Save the project in its source folder: {root}")
        save_project(manifest, req["project"])
        return str(manifest)
    if action == "compile":
        subprocess.run([sys.executable, "-I", HERE / "_xll_compile.py", manifest, req["template"], req["path"]],
                       check=True, stdout=sys.stderr)
        return req["path"]
    if action == "preview":
        path = inside(root, req["name"])
        with path.open("rb") as f:
            data = f.read(32769)
        if b"\0" in data:
            return "Binary file"
        return data[:32768].decode("utf-8-sig", errors="replace") + ("\nâ€¦" if len(data) > 32768 else "")
    if action == "commands":
        return command_names(root, req["entries"])
    if action == "installed":
        return sorted({d.metadata["Name"] + "==" + d.version for d in metadata.distributions()
                       if d.metadata["Name"]}, key=str.casefold)
    if action == "addFiles":
        names = []
        current = manifest.resolve() if manifest else None
        for selected in req["paths"]:
            path = Path(selected).resolve(strict=True)
            if not path.is_relative_to(root) or path == current:
                raise ValueError("Choose source or data files inside the project folder.")
            name = path.relative_to(root).as_posix()
            inside(root, name)
            names.append(name)
        return names
    if action == "wheels":
        return [p.relative_to(root).as_posix() if p.is_relative_to(root) else p.as_posix()
                for p in map(Path, req["paths"])]
    if action in ("image", "importImage"):
        path = Path(req["path"]) if action == "importImage" else inside(root, req["name"])
        data = path.read_bytes()
        if not data.startswith(b"\x89PNG\r\n\x1a\n"):
            raise ValueError("Choose a PNG image.")
        if action == "importImage":
            folder = root / "ribbon-images"
            folder.mkdir(exist_ok=True)
            if folder.resolve().parent != root:
                raise ValueError("The ribbon-images folder must be inside the project.")
            path = folder / (sha256(data).hexdigest()[:16] + ".png")
            try:
                with path.open("xb") as f:
                    f.write(data)
            except FileExistsError:
                if inside(root, path.relative_to(root)).read_bytes() != data:
                    raise ValueError(f"A different image already exists at {path}")
        return dict(name=path.relative_to(root).as_posix(), src="data:image/png;base64," + base64.b64encode(data).decode("ascii"))
    raise ValueError("Unknown designer action")


if __name__ == "__main__":
    try:
        print(json.dumps(handle(json.load(sys.stdin))))
    except Exception as error:
        print(str(error), file=sys.stderr)
        sys.exit(1)
