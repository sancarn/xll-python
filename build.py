"""Build the single 64-bit Windows Excel add-in with Python 3.10+ and VS Build Tools."""
import concurrent.futures
import ctypes
from ctypes import wintypes
from hashlib import sha256
import json
import os
import re
from pathlib import Path
import shutil
import subprocess
import urllib.request
import zipfile

ROOT = Path(__file__).resolve().parent
BUILD = ROOT / ".build"
DIST = ROOT / "dist"
DNA = BUILD / "exceldna.addin/tools/net462"
WEB = BUILD / "microsoft.web.webview2"
INPUTS = json.loads((ROOT / "build-inputs.json").read_text())


def run(*args, **kwargs):
    subprocess.run(list(map(str, args)), check=True, **kwargs)


def prepare(item):
    name, spec = item
    archive = BUILD / "downloads" / name
    archive.parent.mkdir(parents=True, exist_ok=True)
    if not archive.exists() or sha256(archive.read_bytes()).hexdigest() != spec["sha256"]:
        urllib.request.urlretrieve(spec["url"], archive)
    if sha256(archive.read_bytes()).hexdigest() != spec["sha256"]:
        raise ValueError(f"Checksum mismatch for {name}")
    if name.endswith(".nupkg"):
        with zipfile.ZipFile(archive) as source:
            source.extractall(BUILD / archive.stem)
    return archive


def payload(inputs, published=False):
    archive_name = "published-runtime" if published else "runtime"
    inputs = [p for p in inputs if p.suffix == ".msi" and
              (not published or p.name not in ("dev.msi", "test.msi"))]
    runtime = BUILD / (archive_name + ".zip")
    # Regenerate if any pinned input changes. Reuse identical bytes between host-only builds.
    identity = sha256(json.dumps({p.name: INPUTS[p.name] for p in inputs}, sort_keys=True).encode()).hexdigest()
    stamp = BUILD / (archive_name + "-inputs")
    if runtime.exists() and stamp.exists() and stamp.read_text() == identity:
        return runtime
    with zipfile.ZipFile(runtime, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        seen = set()
        for msi in inputs:
            stage = BUILD / "msi" / msi.stem
            stage.mkdir(parents=True, exist_ok=True)
            run("msiexec", "/a", msi, "/qn", "TARGETDIR=" + str(stage))
            for f in sorted(stage.rglob("*")):
                if f.is_file() and f.suffix != ".msi":
                    name = f.relative_to(stage).as_posix()
                    if published and name.lower().startswith(("include/", "libs/")):
                        continue
                    if name not in seen:
                        archive.write(f, name)
                        seen.add(name)
    stamp.write_text(identity)
    return runtime


def embed_archive(xll, name, archive):
    """Store the install archive as an ordinary Win32 RCDATA resource."""
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    begin = kernel.BeginUpdateResourceW
    begin.argtypes = [wintypes.LPCWSTR, wintypes.BOOL]
    begin.restype = wintypes.HANDLE
    update = kernel.UpdateResourceW
    update.argtypes = [wintypes.HANDLE, ctypes.c_void_p, wintypes.LPCWSTR,
                       wintypes.WORD, ctypes.c_void_p, wintypes.DWORD]
    end = kernel.EndUpdateResourceW
    end.argtypes = [wintypes.HANDLE, wintypes.BOOL]
    data = archive.read_bytes()
    handle = begin(str(xll), False)
    if not handle:
        raise ctypes.WinError(ctypes.get_last_error())
    if not update(handle, 10, name, 0, data, len(data)):
        error = ctypes.WinError(ctypes.get_last_error())
        end(handle, True)
        raise error
    if not end(handle, False):
        raise ctypes.WinError(ctypes.get_last_error())


def pack(vs, out, name, dna, runtime, resources=()):
    shutil.copy2(dna, out / (name + ".dna"))
    shutil.copy2(DNA / "ExcelDna64.xll", out / (name + ".xll"))
    run(vs / "MSBuild/Current/Bin/MSBuild.exe", ROOT / "addin/pack.proj",
        "/nologo", "/verbosity:minimal", "/p:AddInName=" + name)
    packed = out / (name + "-packed.xll")
    embed_archive(packed, "PYTHON_RUNTIME", runtime)
    for key, file in resources:
        embed_archive(packed, key, file)
    DIST.mkdir(exist_ok=True)
    artifact = DIST / (name + ".xll")
    os.replace(packed, artifact)
    print(artifact)
    return artifact


def main():
    os.chdir(ROOT)
    with concurrent.futures.ThreadPoolExecutor(5) as pool:
        archives = list(pool.map(prepare, INPUTS.items()))
    runtime = payload(archives)
    from ide_build import bundle
    ide = bundle(ROOT, BUILD, archives)
    installer = Path(os.environ["ProgramFiles(x86)"]) / "Microsoft Visual Studio/Installer/vswhere.exe"
    vs = Path(subprocess.check_output([installer, "-latest", "-products", "*", "-requires",
              "Microsoft.VisualStudio.Component.VC.Tools.x86.x64", "-property", "installationPath"], text=True).strip())
    vcvars = vs / "VC/Auxiliary/Build/vcvars64.bat"
    dev_env = dict(os.environ, PATH=str(installer.parent) + os.pathsep + os.environ["PATH"])
    vcenv = subprocess.check_output(f'call "{vcvars}" >nul && set', shell=True, text=True, errors="replace", env=dev_env)
    env = dict(line.split("=", 1) for line in vcenv.splitlines() if "=" in line and not line.startswith("="))
    out = BUILD / "pack"
    out.mkdir(exist_ok=True)
    dev = BUILD / "msi/dev"
    run(shutil.which("cl.exe", path=next(v for k,v in env.items() if k.lower()=="path")),
        "/nologo", "/O2", "/MT", "/LD", "/I"+str(dev/"include"), ROOT/"native/bridge.c",
        "/link", "/LIBPATH:"+str(dev/"libs"), "/OUT:"+str(out/"bridge.dll"), cwd=out, env=env)
    web_lib = WEB / "lib/net462"
    web_dlls = [web_lib / ("Microsoft.Web.WebView2."+name+".dll") for name in ("Core","WinForms")]
    csc = vs/"MSBuild/Current/Bin/Roslyn/csc.exe"
    run(csc, "/nologo", "/target:winexe", "/platform:x64", "/optimize+", "/out:"+str(out/"Designer.exe"),
        "/r:System.Drawing.dll", "/r:System.Windows.Forms.dll", "/r:System.Web.Extensions.dll",
        *("/r:"+str(f) for f in web_dlls), ROOT/"compiler/DesignerWindow.cs")
    support = out / "support.zip"
    with zipfile.ZipFile(support,"w",zipfile.ZIP_DEFLATED) as z:
        for f in sorted((ROOT/"python").glob("*.py")):
            z.write(f,f.name)
        wheel = next(p for p in archives if p.name.startswith("debugpy-"))
        with zipfile.ZipFile(wheel) as src:
            for item in src.infolist():
                if not item.is_dir():
                    z.writestr(item.filename, src.read(item))
        for folder in ("repl", "packages", "compiler"):
            for f in sorted((ROOT/"addin"/folder).iterdir()):
                z.write(f, folder + "/" + f.name)
        for f in (out/"Designer.exe", *web_dlls, out/"bridge.dll", WEB/"build/native/x64/WebView2Loader.dll"):
            z.write(f,f.name)
        z.write(ROOT/"examples/functions.py","examples/functions.py")
        z.write(ROOT/"examples/LICENSE", "examples/LICENSE")
        for name in ("LICENSE", "LICENSE-PACKAGE-EXCEPTIONS.txt"):
            z.write(ROOT/name, name)
        for f in sorted((ROOT/"notices").glob("*")):
            z.write(f,"notices/"+f.name)
        for name in ("LICENSE.txt", "NOTICE.txt"):
            z.write(WEB/name, "notices/WebView2-"+name)
    manifest = out/"versions.txt"
    manifest.write_text("\n".join(sha256(f.read_bytes()).hexdigest() for f in (runtime,support,ide)))
    refs = ["System.dll","System.Core.dll","System.Drawing.dll","System.Windows.Forms.dll",
            "System.Web.Extensions.dll","System.IO.Compression.dll","System.Net.Http.dll","Microsoft.CSharp.dll"]
    release = os.environ.get("PYTHONEXCEL_RELEASE_TAG", "v0.2.0")
    if not re.fullmatch(r"v\d+\.\d+\.\d+", release):
        raise ValueError("Expected a release tag such as v0.2.1")
    (out/"release.txt").write_text(release, encoding="ascii")
    run(csc, "/nologo", "/target:winexe", "/optimize+", "/out:"+str(out/"updater.exe"),
        "/r:System.Web.Extensions.dll", "/r:System.Windows.Forms.dll", *sorted((ROOT/"updater").glob("*.cs")))
    run(csc,"/nologo","/target:library","/optimize+","/platform:x64","/out:"+str(out/"PythonExcel.dll"),
        *("/r:"+x for x in refs), "/r:"+str(DNA/"ExcelDna.Integration.dll"),
        *("/r:"+str(f) for f in web_dlls),
        "/resource:"+str(support)+",support.zip", "/resource:"+str(manifest)+",versions.txt",
        "/resource:"+str(out/"updater.exe")+",updater.exe", "/resource:"+str(out/"release.txt")+",release.txt",
        ROOT/"updater/RestartPlan.cs", *sorted((ROOT/"addin").glob("*.cs")))
    for f in web_dlls:
        shutil.copy2(f,out/f.name)
    lean_runtime = payload(archives, published=True)
    lean_support = out / "published-support.zip"
    with zipfile.ZipFile(lean_support, "w", zipfile.ZIP_DEFLATED) as z:
        for module in ("_xll_worker", "_xll_functions", "_xll_application", "_xll_streams", "xll_python"):
            z.write(ROOT / "python" / (module + ".py"), module + ".py")
        for name in ("LICENSE", "LICENSE-PACKAGE-EXCEPTIONS.txt", "notices/Excel-DNA.txt", "notices/Third-party-sources.md"):
            z.write(ROOT / name, name)
    lean_versions = out / "published-versions.txt"
    lean_versions.write_text("\n".join(sha256(p.read_bytes()).hexdigest() for p in (lean_runtime, lean_support)))
    run(csc, "/nologo", "/target:library", "/optimize+", "/platform:x64", "/define:RUNTIME_ONLY",
        "/out:" + str(out / "PythonExcelRuntime.dll"), *("/r:" + x for x in refs),
        "/r:" + str(DNA / "ExcelDna.Integration.dll"),
        "/resource:" + str(lean_support) + ",support.zip", "/resource:" + str(lean_versions) + ",versions.txt",
        ROOT / "addin/runtime/Ribbon.cs", *(ROOT / "addin" / (name + ".cs") for name in
        ("Runtime", "Published", "Functions", "ExcelObjects", "LiveFunction", "ProcessJob")))
    template = pack(vs, out, "PythonExcelRuntime", ROOT / "addin/runtime/PythonExcelRuntime.dna", lean_runtime)
    pin = out / "runtime-template.json"
    pin.write_text(json.dumps(dict(tag=release, sha256=sha256(template.read_bytes()).hexdigest())), encoding="ascii")
    pack(vs, out, "xll-python", ROOT / "addin/PythonExcel.dna", runtime, [("EXCEL_IDE", ide), ("RUNTIME_TEMPLATE", pin)])


if __name__ == "__main__":
    main()
