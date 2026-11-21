# Third-party software in PythonExcel.xll

The following source versions correspond to the components bundled in this release.

- **Excel-DNA 1.9.0** supplies the Excel add-in loader and integration library. Its zlib licence is in `Excel-DNA.txt`. [Excel-DNA source](https://github.com/Excel-DNA/ExcelDna/tree/v1.9.0-Release)
- **CPython 3.13.12** supplies the Windows interpreter and standard library. Its `LICENSE.txt` is included at the root of the extracted Python runtime. The runtime also includes components with their own terms, including Tcl/Tk and the packages vendored in pip. [CPython source](https://github.com/python/cpython/tree/v3.13.12)
- **Eclipse Theia 1.77.0** supplies the IDE framework. We distribute it under the Eclipse Public License 2.0. The `Theia-*` files reproduce upstream's licences and notice. The linked tag contains the framework source used by this build. The Excel integration and the scripts that build the workbench are in the XLL Python release source. [Theia source](https://github.com/eclipse-theia/theia/tree/v1.77.0)

The extracted support directory also includes the WebView2 SDK licence and notice and the debugpy wheel's licence files. The extracted IDE contains the Node.js licence, npm package notices, and the licences shipped with the Pyright and Python grammar extensions. The versioned build inputs are recorded in `build-inputs.json` and `ide/package-lock.json` in the XLL Python release source.
