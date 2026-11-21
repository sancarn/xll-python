using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

internal sealed class DesignerWindow : Form
{
    const string Host = "xll-designer.example";
    const string Page = "https://" + Host + "/index.html";
    readonly WebView2 web = new WebView2 { Dock = DockStyle.Fill };
    readonly JavaScriptSerializer json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
    readonly string python, template;
    string root, manifest;
    bool ready, busy, closing;

    [STAThread]
    static void Main(string[] args)
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new DesignerWindow(args[0], args[1], args[2]));
    }

    DesignerWindow(string python, string template, string profile)
    {
        this.python = python;
        this.template = template;
        Text = "Compile to XLL";
        ClientSize = new Size(1000, 720);
        MinimumSize = new Size(660, 480);
        StartPosition = FormStartPosition.CenterScreen;
        Controls.Add(web);
        Shown += async (sender, args) => {
            try
            {
                CoreWebView2Environment.SetLoaderDllFolderPath(AppDomain.CurrentDomain.BaseDirectory);
                web.CreationProperties = new CoreWebView2CreationProperties { UserDataFolder = profile };
                await web.EnsureCoreWebView2Async();
                if (IsDisposed) return;
                var core = web.CoreWebView2;
                core.Settings.AreHostObjectsAllowed = false;
                core.Settings.AreBrowserAcceleratorKeysEnabled = false;
                core.Settings.AreDefaultContextMenusEnabled = false;
                core.SetVirtualHostNameToFolderMapping(Host, Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "compiler"),
                    CoreWebView2HostResourceAccessKind.Deny);
                core.NavigationStarting += (s,e) => e.Cancel = ready || e.Uri != Page;
                core.NewWindowRequested += (s,e) => e.Handled = true;
                core.WebMessageReceived += Receive;
                core.ProcessFailed += (s,e) => {
                    if (e.ProcessFailedKind == CoreWebView2ProcessFailedKind.RenderProcessExited ||
                        e.ProcessFailedKind == CoreWebView2ProcessFailedKind.BrowserProcessExited)
                        Fail("The designer stopped responding. Reopen Compile to XLL.");
                };
                core.NavigationCompleted += (s,e) => { if (!e.IsSuccess && !ready) Fail("Could not load the project designer."); };
                core.Navigate(Page);
            }
            catch (Exception error) { if (!IsDisposed) Fail(error.Message); }
        };
        FormClosing += (s,e) => {
            if (closing || !ready) return;
            e.Cancel = true;
            if (!busy) web.CoreWebView2.PostWebMessageAsJson("{\"event\":\"close\"}");
        };
    }

    void Fail(string message)
    {
        ready = false;
        web.Hide();
        Controls.Add(new Label { Text = message, Dock = DockStyle.Fill, Padding = new Padding(20) });
    }

    string PickFile(bool save, string title, string filter, string initial = "")
    {
        using (FileDialog dialog = save ? (FileDialog)new SaveFileDialog() : new OpenFileDialog())
        {
            dialog.Title = title;
            dialog.Filter = filter;
            dialog.InitialDirectory = root;
            dialog.FileName = initial;
            dialog.RestoreDirectory = true;
            return dialog.ShowDialog(this) == DialogResult.OK ? dialog.FileName : null;
        }
    }

    async void Receive(object sender, CoreWebView2WebMessageReceivedEventArgs args)
    {
        if (args.Source != Page || web.CoreWebView2.Source != Page) return;
        var req = json.Deserialize<Dictionary<string,object>>(args.WebMessageAsJson);
        var id = req["id"];
        if (busy) return;
        busy = true;
        try
        {
            string action = (string)req["action"];
            if (action == "ready") ready = true;
            else if (action == "close") { closing = true; Close(); return; }
            else { Reply(new { id, result = await Dispatch(action, req, args) }); return; }
            Reply(new { id, result = (object)null });
        }
        catch (Exception error) { Reply(new { id, error = error.Message }); }
        finally { busy = false; }
    }

    async Task<object> Dispatch(string action, Dictionary<string,object> req, CoreWebView2WebMessageReceivedEventArgs args)
    {
        // Commit a folder or filename to this session only after the operation succeeds.
        string folder = root, file = manifest;
        switch (action)
        {
            case "new":
                if (args.AdditionalObjects?.Count > 0)
                {
                    folder = (args.AdditionalObjects.Single() as CoreWebView2File)?.Path;
                    if (!Directory.Exists(folder)) throw new IOException("Drop one project folder.");
                }
                else using (var picker = new FolderBrowserDialog { Description = "Select your project folder" })
                {
                    if (picker.ShowDialog(this) != DialogResult.OK) return null;
                    folder = picker.SelectedPath;
                }
                file = null;
                break;
            case "open":
                file = PickFile(false, "Open project", "XLL projects (*.json)|*.json");
                if (file == null) return null;
                folder = Path.GetDirectoryName(file);
                break;
            case "save":
                if (file == null)
                {
                    file = PickFile(true, "Save project in its source folder", "XLL projects|*.xllproject.json", "project.xllproject.json");
                    if (file == null) return null;
                }
                break;
            case "compile":
                string name = string.Concat(((string)req["name"]).Split(Path.GetInvalidFileNameChars()));
                req["path"] = PickFile(true, "Save compiled add-in", "Excel add-ins|*.xll", name + ".xll");
                if (req["path"] == null) return null;
                req["template"] = template;
                break;
            case "importImage":
                req["path"] = PickFile(false, "Import ribbon image", "PNG images|*.png");
                if (req["path"] == null) return null;
                break;
            case "addFiles": case "wheels":
                using (var picker = new OpenFileDialog {
                    Title = action == "addFiles" ? "Add project files" : "Add local wheels",
                    Filter = action == "addFiles" ? "All files|*.*" : "Python wheels|*.whl",
                    Multiselect = true, InitialDirectory = root
                })
                {
                    if (picker.ShowDialog(this) != DialogResult.OK) return null;
                    req["paths"] = picker.FileNames;
                }
                break;
            case "refresh": case "preview": case "installed": case "commands": case "image":
                break;
            default: throw new InvalidOperationException("Unknown project action.");
        }
        req["root"] = folder;
        req["manifest"] = file;
        var result = await Worker(req);
        root = folder;
        manifest = file;
        return result;
    }

    void Reply(object value)
    {
        if (!IsDisposed) web.CoreWebView2.PostWebMessageAsJson(json.Serialize(value));
    }

    async Task<object> Worker(Dictionary<string,object> request)
    {
        var info = new ProcessStartInfo(python, "-I -X utf8 _xll_designer.py") {
            WorkingDirectory = AppDomain.CurrentDomain.BaseDirectory, UseShellExecute = false, CreateNoWindow = true,
            RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true,
            StandardErrorEncoding = Encoding.UTF8, StandardOutputEncoding = Encoding.UTF8
        };
        using (var p = Process.Start(info))
        {
            var stdout = p.StandardOutput.ReadToEndAsync();
            var stderr = p.StandardError.ReadToEndAsync();
            var bytes = Encoding.UTF8.GetBytes(json.Serialize(request));
            await p.StandardInput.BaseStream.WriteAsync(bytes, 0, bytes.Length);
            p.StandardInput.Close();
            await Task.WhenAll(stdout, stderr);
            p.WaitForExit();
            if (p.ExitCode != 0) {
                string detail = stderr.Result.Trim();
                throw new IOException(detail.Substring(Math.Max(0, detail.Length - 4000)));
            }
            return json.DeserializeObject(stdout.Result);
        }
    }
}
