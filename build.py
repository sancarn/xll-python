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


def payload(inputs):
    runtime = BUILD / "runtime.zip"
    # Regenerate if any pinned input changes. Reuse identical bytes between host-only builds.
    identity = sha256(json.dumps({k:v for k,v in INPUTS.items() if k.endswith(".msi")}, sort_keys=True).encode()).hexdigest()
    stamp = BUILD / "runtime-inputs"
    if runtime.exists() and stamp.exists() and stamp.read_text() == identity:
        return runtime
    with zipfile.ZipFile(runtime, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        seen = set()
        for msi in inputs:
            if msi.suffix != ".msi":
                continue
            stage = BUILD / "msi" / msi.stem
            stage.mkdir(parents=True, exist_ok=True)
            run("msiexec", "/a", msi, "/qn", "TARGETDIR=" + str(stage))
            for f in sorted(stage.rglob("*")):
                if f.is_file() and f.suffix != ".msi":
                    name = f.relative_to(stage).as_posix()
                    if name not in seen:
                        archive.write(f, name)
                        seen.add(name)
    stamp.write_text(identity)
    return runtime


def embed_runtime(xll, runtime):
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
    data = runtime.read_bytes()
    handle = begin(str(xll), False)
    if not handle:
        raise ctypes.WinError(ctypes.get_last_error())
    if not update(handle, 10, "PYTHON_RUNTIME", 0, data, len(data)):
        error = ctypes.WinError(ctypes.get_last_error())
        end(handle, True)
        raise error
    if not end(handle, False):
        raise ctypes.WinError(ctypes.get_last_error())


def main():
    os.chdir(ROOT)
    with concurrent.futures.ThreadPoolExecutor(5) as pool:
        archives = list(pool.map(prepare, INPUTS.items()))
    runtime = payload(archives)
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
    support = out / "support.zip"
    with zipfile.ZipFile(support,"w",zipfile.ZIP_DEFLATED) as z:
        for f in sorted((ROOT/"python").glob("*.py")):
            z.write(f,f.name)
        for f in (out/"bridge.dll", WEB/"build/native/x64/WebView2Loader.dll"):
            z.write(f,f.name)
        z.write(ROOT/"examples/functions.py","examples/functions.py")
        for f in sorted((ROOT/"notices").glob("*")):
            z.write(f,"notices/"+f.name)
        for name in ("LICENSE.txt", "NOTICE.txt"):
            z.write(WEB/name, "notices/WebView2-"+name)
    manifest = out/"versions.txt"
    manifest.write_text("\n".join(sha256(f.read_bytes()).hexdigest() for f in (runtime,support)))
    refs = ["System.dll","System.Core.dll","System.Drawing.dll","System.Windows.Forms.dll",
            "System.Web.Extensions.dll","System.IO.Compression.dll","System.Net.Http.dll","Microsoft.CSharp.dll"]
    csc = vs/"MSBuild/Current/Bin/Roslyn/csc.exe"
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
    shutil.copy2(ROOT/"addin/PythonExcel.dna",out/"PythonExcel.dna")
    shutil.copy2(DNA/"ExcelDna64.xll",out/"PythonExcel.xll")
    DIST.mkdir(exist_ok=True)
    run(vs/"MSBuild/Current/Bin/MSBuild.exe", ROOT/"addin/pack.proj", "/nologo", "/verbosity:minimal")
    artifact = DIST / "PythonExcel.xll"
    embed_runtime(out/"PythonExcel-packed.xll", runtime)
    os.replace(out/"PythonExcel-packed.xll", artifact)
    print(artifact)


if __name__ == "__main__":
    main()

