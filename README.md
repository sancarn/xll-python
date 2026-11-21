<img src="assets/icon.svg" width="64" height="64" alt="">

# Python for desktop Excel

[Download the latest 64-bit XLL](https://github.com/sancarn/xll-python/releases/latest/download/PythonExcel.xll). It contains CPython 3.13.12 for Windows x64, its standard library, the native DLLs, Tcl/Tk, development headers and import libraries, pip’s offline bootstrap wheel, and the Excel add-in. Worksheet functions and scripts run in the Excel process. Each pip command runs in a separate process using the same packaged interpreter.

## Load the add-in

1. Save `PythonExcel.xll` to a permanent local folder. If Windows marked the downloaded file as coming from the internet, review its source and use **Properties → Unblock** if appropriate. This build is unsigned.
2. In **64-bit desktop Excel for Windows**, open **File → Options → Add-ins → Manage Excel Add-ins → Go → Browse**. Select the XLL and leave it checked.
3. Open the **Python** ribbon tab. The first operation that needs Python extracts the embedded payload under `%LOCALAPPDATA%\XllPython` and bootstraps pip from the bundled wheel. Installation works offline. Later launches reuse the installed files. With no configured folders, Excel opens the ribbon and initializes Python on first use. The add-in uses Windows’ .NET Framework 4.8.

**Run Script** opens a Windows file picker. Select a Python file to run it as `__main__`. Its folder is on the import path while it runs. A window displays the script’s output or traceback.

**Check for Updates** downloads the latest release from this repository and verifies its SHA-256 digest. Before restarting this Excel instance, the updater gives you three choices.

- **Save all** saves your changes. Excel asks where to save each new workbook or edited read-only workbook. Canceling Save As stops the update and leaves the workbooks open. Any saves already completed are kept.
- **Don't save any** closes every workbook in this instance and discards unsaved changes. After the restart, Excel reopens the last saved versions. Workbooks that were never saved are discarded.
- **Cancel** stops the update and leaves all workbooks open with their edits intact.

After Excel exits, a separate helper replaces the XLL at its existing path and reopens the saved workbooks. The add-in folder must be writable. If another Excel instance locks the XLL, the helper reopens the workbooks with the existing version.

**Manage Packages** opens a WebView2 window for the add-in’s Python environment.

- **Browse** suggests packages for Excel and searches PyPI by project name. The first search downloads the name index. Later searches reuse the local cache. Choose **Add** to queue an install. Its row shows progress and a tick once the package is installed. If an install fails, choose **View log** for details or **Retry**.
- **Installed** lists packages and versions, including dependencies. The list refreshes after each command and works offline.
- **Advanced** accepts pip commands such as `list`, `install "C:\downloads\some_package.whl"`, and `uninstall -y openpyxl`, with or without the leading `pip`. Each command has its own live log and exit code. You can submit more commands while one runs. Browse and Advanced share a queue that runs in order and stays available when you reopen the window in the same Excel session.

Installed packages persist for that runtime. Restart Excel after replacing an imported package or installing a package that uses a startup `.pth` file.

## Worksheet functions

In **Python → Startup Folders → Worksheet functions**, add a folder containing your Python files. Choose **Refresh** to save the folder list, reload the Python files, and recalculate open workbooks. **Save** applies the folder list at the next Excel startup. The add-in scans top-level `.py` files, skipping names beginning with an underscore. A configured folder runs Python at Excel startup, so choose folders whose contents you trust. After editing a function or a helper module in a startup folder, choose **Refresh** again. A source or import error preserves the last working definitions. Functions are available in every open workbook.

To try the [working examples](examples/functions.py), add the repository’s `examples` folder. The XLL also extracts a copy under `%LOCALAPPDATA%\XllPython\support-…\examples`; copy it to your own editable folder before customizing it.

```python
import asyncio
from typing import Annotated
from xll_python import Function, function

@function(name="XLLPY.scale", category="Finance", thread_safe=True)
def scale(value: Annotated[float, "Number to scale"], factor: float = 2) -> float:
    """Multiply a number by a scale factor."""
    return value * factor

async def delayed(value: float, seconds: float = 2) -> Annotated[float, Function(name="XLLPY.delayed")]:
    await asyncio.sleep(seconds)
    return value
```

Enter `=XLLPY.scale(21)` for **42**. With `=XLLPY.delayed(123,5)`, Excel displays **Waiting…** while Python waits five seconds on a background thread. You can edit other cells during the wait. Excel receives the result through real-time data (RTD) and updates the cell to **123** along with its dependent formulas. A regular `def` runs synchronously by default. Use `asynchronous=True` to offload a blocking regular function. Decorators and `Annotated` return metadata accept the same registration options.

An exported generator publishes each `yield` as a live worksheet value. Both `def` and `async def` generators run in the background. Excel shows **Waiting…** until the first yield and displays the latest value at its RTD refresh interval. A finite generator keeps its last value, or supplies an empty result if it never yields.

By default, `generator_instances="one"` shares one generator across cells with the same function and arguments. Use `generator_instances="many"` for a separate instance per calling cell. Try filling a column with `=XLLPY.live_random()` for independent sequences that generate a number every two seconds. Recalculation keeps each subscription. Removing its last formula closes the generator and runs its `finally` block. A synchronous generator blocked in a call such as `time.sleep` can finish that call before closing. **Refresh** starts new subscriptions using the reloaded definitions.

- `name`, `category`, `description` and `help_topic` control registration and Function Wizard metadata. The docstring supplies the default description. An `Annotated` parameter can supply its argument description.
- `volatile`, `thread_safe`, `macro_type`, `hidden` and `cluster_safe` expose Excel registration flags. Use flags that match the implementation. A cluster worker would also need the add-in, functions and their dependencies.
- Give functions a fixed positional signature. Omitted arguments use Python defaults. Empty cells become `None`; worksheet errors become `xll_python.ExcelError`. Python exceptions return `#VALUE!` and write a traceback to `%LOCALAPPDATA%\XllPython\python-excel.log`.
- Ranges arrive as lists of rows. A rectangular return value spills in dynamic-array Excel. A flat list returns a column. Annotating a parameter with `xll_python.Reference` requests its sheet identifier and zero-based area coordinates. Date annotations use OLE Automation serial dates.
- On **Startup Folders → Package links**, add existing `site-packages` directories or package parent folders. Native extensions need a Windows x64 CPython 3.13-compatible ABI. To export functions defined in an installed package, import them in a startup file and list the decorated functions in `__excel_exports__`, for example `from mypackage import price` followed by `__excel_exports__ = [price]`.

## Python console

Choose **Python → Console** to open an interactive session in Excel's embedded interpreter. The console uses the Microsoft Edge WebView2 Runtime and works offline. Installed packages and configured import paths are available in the session.

```python
sales = {"North": {"units": [12, 18]}, "South": {"units": [9]}}
print("Sales by region", sales)
```

Expand the dictionary beside **Sales by region**, then expand a region and its list. Entering `sales` by itself also produces an expandable result. The inspector reads live objects when you expand them and pages through large collections. Standard output, standard error and tracebacks from commands appear in the transcript as they run.

Press **Enter** to run a command or **Shift+Enter** to insert a line. Incomplete statements continue on the next line. Type a name or a dot to see completions, or press **Ctrl+Space** to request them. Use the arrow keys and **Tab** to choose a suggestion. With suggestions closed, **Up** and **Down** recall previous commands.

Variables persist while the console window is open. **Clear output** clears the transcript and releases inspector references while keeping your variables. Reopening a closed window starts a fresh session. If you close the window during a command, that command can finish in the background.

To test a function in a worksheet, define it in the console:

```python
def double(x):
    return 2 * x
```

Enter `21` in A1 and `=XLLPY.console.double(A1)` in B1. B1 shows **42**. Redefine `double` in the console and Excel recalculates the existing formula. The function can use console variables. For a plain function, press **F9** after changing a variable it reads.

Public functions defined in the console with fixed positional parameters are available in every workbook in that Excel instance. Use a leading underscore for a private helper. A startup function with the same Excel name takes precedence. Deleting the definition or closing the console removes the worksheet function. To keep a function across Excel sessions, save it in a startup folder with `@function(name="XLLPY.console.double")` to keep the same formula.

## VS Code web IDE

**Open IDE** hosts the live **vscode.dev** editor in WebView2. This feature needs internet access to the editor and the Microsoft Edge WebView2 Runtime. The interpreter, worksheet registration, script runner and pip bootstrap work independently of the editor.

1. Use VS Code’s **File → Open File** to select a local Python file. The native toolbar displays the selected path. If you opened a folder or switched editor tabs, use **Choose script** to select the execution target shown in the toolbar.
2. Edit and **save** in VS Code, then choose **Run saved script**. The output pane shows the result from Excel’s embedded interpreter.
3. To debug a script or worksheet function, choose its source file, set a **Breakpoint line**, and select **Arm breakpoint**. Use line **0** to pause on its first executable line. Run the script, evaluate a registered function, or recalculate its worksheet cell. The debugger shows the source line and local variables. Use **Step**, **Continue** or **Detach**. Closing the IDE releases a paused call.
4. Enter an expression such as `=XLLPY.scale(11)` in **Excel formula** and select **Evaluate**. Excel evaluates it in the active workbook. The IDE displays the result without writing to a cell. Use **Startup Folders → Refresh** to reload changes to startup function definitions.

The debugger controls belong to the XLL’s IDE window. They can inspect the live interpreter even when a synchronous worksheet call has paused Excel’s calculation thread.

## Build and verify

Each push to `main` runs a Windows GitHub Actions build. The workflow checks the packaged runtime and restart helper, then publishes the XLL with a versioned tag and SHA-256 digest. Users on v0.1.0 can install the latest release manually once to get the update button.

On Windows x64, install Python 3.10 or later and Visual Studio Build Tools with desktop C++ tools, MSBuild, the C# compiler and .NET Framework 4.8 targeting support. From the repository root run:

```powershell
python build.py
```

The build verifies the SHA-256 of each pinned upstream input in `build-inputs.json`, stages the official CPython MSI payloads without installing Python system-wide, compiles the native embedding bridge and ribbon, and packs `dist\PythonExcel.xll`. Build downloads stay in `.build`. The extracted support directory contains third-party notices, and the runtime contains CPython’s license.

To exercise the finished file in a disposable Excel instance, install `pywin32` in the **test driver’s** Python and run `python tests/excel_smoke.py`. For the ribbon and IDE checks, also install `pywinauto`, `psutil` and `websocket-client`, set `$env:XLL_WEB_TEST = '1'`, and run `python tests/excel_smoke.py --ui`. The UI checks open temporary test windows on the desktop. The driver scopes its workbook, package installs and configuration to a temporary profile via `XLL_PYTHON_HOME`.
