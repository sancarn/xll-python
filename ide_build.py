"""Bundle the local Theia application and its built-in extensions."""
import hashlib
import json
import os
import subprocess
import zipfile


def bundle(root, build, archives):
    ide = root / "ide"
    node_zip = next(p for p in archives if p.name.startswith("node-v"))
    node = build / node_zip.stem / "node.exe"
    if not node.exists():
        with zipfile.ZipFile(node_zip) as src:
            src.extractall(build)
    env = dict(os.environ, PATH=str(node.parent) + os.pathsep + os.environ["PATH"])
    modules = ide / "node_modules"
    npm = [node, node.parent / "node_modules/npm/bin/npm-cli.js"]
    lock = hashlib.sha256((ide / "package-lock.json").read_bytes()).hexdigest()
    stamp = modules / ".xll-build-input"
    if not stamp.exists() or stamp.read_text() != lock:
        subprocess.run([*npm, "ci", "--ignore-scripts", "--no-audit", "--no-fund"], cwd=ide, env=env, check=True)
        subprocess.run([*npm, "rebuild", "@vscode/windows-ca-certs", "keytar", "drivelist"], cwd=ide, env=env, check=True)
        stamp.write_text(lock)
    subprocess.run([node, modules / "@theia/cli/bin/theia.js", "build"], cwd=ide, env=env, check=True)

    dest = build / "ide.zip"
    with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        z.write(node, "node.exe")
        z.write(node.parent / "LICENSE", "notices/Node-LICENSE")
        z.write(ide / "package.json", "package.json")
        for file in sorted((ide / "lib").rglob("*")):
            if file.is_file() and file.suffix not in (".gz", ".map"):
                z.write(file, file.relative_to(ide))
        for file in (ide / "plugin").iterdir():
            z.write(file, "plugins/excel-python/" + file.name)
        for vsix in (p for p in archives if p.suffix == ".vsix"):
            with zipfile.ZipFile(vsix) as src:
                for item in src.infolist():
                    if item.filename.startswith("extension/") and not item.is_dir() and not item.filename.endswith(".map"):
                        z.writestr("plugins/" + vsix.stem + "/" + item.filename[10:], src.read(item))
        packages = json.loads((ide / "package-lock.json").read_text())["packages"]
        for name, spec in packages.items():
            if not name or spec.get("dev") or spec.get("link"):
                continue
            directory = ide / name
            if directory.is_dir():
                for file in directory.iterdir():
                    if file.is_file() and file.name.lower().startswith(("license", "licence", "notice", "copying")):
                        z.write(file, "notices/" + name + "/" + file.name)
    return dest
