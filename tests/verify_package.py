"""Check the CI-built PE resource and packaged support files without desktop Excel."""
import ctypes
from ctypes import wintypes
import hashlib
from pathlib import Path
import zipfile


def main():
    root = Path(__file__).resolve().parents[1]
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.LoadLibraryExW.argtypes = [wintypes.LPCWSTR, wintypes.HANDLE, wintypes.DWORD]
    kernel.LoadLibraryExW.restype = wintypes.HANDLE
    kernel.FindResourceW.argtypes = [wintypes.HANDLE, wintypes.LPCWSTR, wintypes.LPCWSTR]
    kernel.FindResourceW.restype = wintypes.HANDLE
    kernel.SizeofResource.argtypes = [wintypes.HANDLE, wintypes.HANDLE]
    kernel.LoadResource.argtypes = [wintypes.HANDLE, wintypes.HANDLE]
    kernel.LoadResource.restype = wintypes.HANDLE
    kernel.LockResource.argtypes = [wintypes.HANDLE]
    kernel.LockResource.restype = wintypes.LPVOID
    kernel.FreeLibrary.argtypes = [wintypes.HANDLE]
    module = kernel.LoadLibraryExW(str(root/"dist/PythonExcel.xll"), None, 2)
    if not module:
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        for name, kind, source in [
            ("PYTHON_RUNTIME", ctypes.cast(10, wintypes.LPCWSTR), root/".build/runtime.zip"),
            ("EXCEL_IDE", ctypes.cast(10, wintypes.LPCWSTR), root/".build/ide.zip"),
            ("PYTHONEXCEL", "ASSEMBLY", root/".build/pack/PythonExcel.dll"),
        ]:
            entry = kernel.FindResourceW(module, name, kind)
            if not entry:
                raise ctypes.WinError(ctypes.get_last_error())
            pointer = kernel.LockResource(kernel.LoadResource(module, entry))
            assert pointer
            data = ctypes.string_at(pointer, kernel.SizeofResource(module, entry))
            assert data == source.read_bytes(), name
            if name == "PYTHONEXCEL":
                assert (root/".build/pack/updater.exe").read_bytes() in data
                support = root/".build/pack/support.zip"
                assert support.read_bytes() in data
                with zipfile.ZipFile(support) as archive:
                    for path in ("LICENSE", "LICENSE-PACKAGE-EXCEPTIONS.txt", "examples/LICENSE",
                                 "notices/Excel-DNA.txt", "notices/Theia-LICENSE-EPL",
                                 "notices/Theia-NOTICE.md", "notices/Third-party-sources.md"):
                        assert archive.read(path) == (root/path).read_bytes(), path
    finally:
        kernel.FreeLibrary(module)
    with zipfile.ZipFile(root/".build/runtime.zip") as archive:
        assert archive.testzip() is None
        assert b"PYTHON SOFTWARE FOUNDATION LICENSE" in archive.read("LICENSE.txt")
    print("Verified single-XLL package", hashlib.sha256((root/"dist/PythonExcel.xll").read_bytes()).hexdigest())


if __name__ == "__main__":
    main()
