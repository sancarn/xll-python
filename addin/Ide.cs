using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace PythonExcel
{
    internal sealed class IdeWindow : Form
    {
        static readonly WebWindowHost host = new WebWindowHost("Excel Python IDE", () => new IdeWindow());
        readonly WebView2 web = new WebView2 { Dock = DockStyle.Fill };
        readonly string token = Guid.NewGuid().ToString("N") + Guid.NewGuid().ToString("N");
        readonly int pid = Process.GetCurrentProcess().Id;
        readonly TaskCompletionSource<int> listening = new TaskCompletionSource<int>();
        Process backend;
        ProcessJob job;
        Uri origin;
        bool closed, ready, closing, allowClose;
        readonly Stopwatch closeWait = new Stopwatch();
        string version;
        int refreshVersion;

        internal static void Open() => host.Open();
        internal static void Shutdown() => host.Shutdown();

        IdeWindow()
        {
            Text = "Python IDE — Excel";
            ClientSize = new Size(1140, 780);
            MinimumSize = new Size(760, 500);
            Controls.Add(web);
            Shown += async (s, e) => {
                try { await Start(); }
                catch (Exception error) {
                    if (!closed) {
                        Runtime.Log(error.ToString());
                        MessageBox.Show(this, error.Message, "Excel Python IDE");
                        Close();
                    }
                }
            };
            FormClosing += (s, e) => {
                if (host.IsClosing || !ready || allowClose || e.CloseReason != CloseReason.UserClosing) return;
                if (closing && closeWait.Elapsed >= TimeSpan.FromSeconds(5) &&
                    MessageBox.Show(this, "Close the IDE now and discard any unsaved edits?", "Close Excel Python IDE",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                    return;
                e.Cancel = true;
                if (closing) return;
                closeWait.Restart();
                closing = true;
                web.CoreWebView2.PostWebMessageAsJson(@"{""type"":""closing""}");
            };
            FormClosed += (s, e) => {
                closed = true;
                listening.TrySetCanceled();
                job?.Dispose();
                backend?.Dispose();
                // Detach also when the renderer has exited and its cookie manager is unavailable.
                Task.Run(() => {
                    try { Runtime.Request(new { op = "ideDisconnect", session = token }); }
                    catch (Exception error) { Runtime.Log(error.ToString()); }
                });
                try {
                    if (origin != null && web.CoreWebView2 != null)
                        web.CoreWebView2.CookieManager.DeleteCookies("xll_" + origin.Port, origin.AbsoluteUri);
                } catch (Exception error) { Runtime.Log(error.ToString()); }
            };
        }

        string WriteWorkspaceDefaults(IDictionary environment)
        {
            var settings = Runtime.LoadSettings();
            var roots = settings.folders.Concat(settings.packages).Distinct(StringComparer.OrdinalIgnoreCase);
            var paths = ((IList)environment["paths"]).Cast<string>().Where(Directory.Exists);
            var preferences = new Dictionary<string, object> {
                { "python.pythonPath", Runtime.Python }, { "python.defaultInterpreterPath", Runtime.Python },
                { "python.analysis.extraPaths", paths.ToArray() }, { "telemetry.telemetryLevel", "off" }
            };
            var workspace = new {
                folders = roots.Select(path => new { path }).ToArray(),
                settings = preferences,
                launch = new { version = "0.2.0", configurations = new object[] {
                    new { type = "excel-python", request = "attach", name = "Debug Current File in Excel", currentFile = true },
                    new { type = "excel-python", request = "attach", name = "Debug Worksheet Functions" }
                }}
            };
            byte[] content = Encoding.UTF8.GetBytes(Runtime.Json.Serialize(workspace));
            string key;
            using (var hash = SHA256.Create())
                key = BitConverter.ToString(hash.ComputeHash(content), 0, 12).Replace("-", "");
            string dir = Path.Combine(Runtime.Root, "IDE", "workspaces", key);
            Directory.CreateDirectory(dir);
            string file = Path.Combine(dir, "defaults.json");
            if (!File.Exists(file)) {
                string temp = file + "." + Guid.NewGuid().ToString("N");
                File.WriteAllBytes(temp, content);
                try { File.Move(temp, file); }
                catch (IOException) when (File.Exists(file)) { File.Delete(temp); }
            }
            return file;
        }

        async Task Start()
        {
            WebWindowHost.Configure(web, "IdeWebView2");
            await web.EnsureCoreWebView2Async();
            if (closed) return;
            web.CoreWebView2.NavigateToString("<html><body style='font:16px Segoe UI;padding:2em'>Opening Excel Python IDE…</body></html>");
            string dir = await Task.Run(Runtime.InstallIde);
            if (closed) return;
            var env = (IDictionary)await Task.Run(() => Runtime.Request(new { op = "ideEnvironment" }));
            if (closed) return;
            version = (string)env["version"];
            string defaults = WriteWorkspaceDefaults(env);
            string worksheet = Path.Combine(Runtime.Root, "IDE", "workspaces", "Excel Python.theia-workspace");
            var info = new ProcessStartInfo(Path.Combine(dir, "node.exe"),
                "lib/backend/main.js --hostname 127.0.0.1 --port 0 \"" + worksheet + "\"") {
                WorkingDirectory = dir, UseShellExecute = false, CreateNoWindow = true,
                RedirectStandardOutput = true, RedirectStandardError = true
            };
            string hostname = "xll-" + Guid.NewGuid().ToString("N") + ".localhost";
            info.EnvironmentVariables["XLL_IDE_DEFAULTS"] = defaults;
            info.EnvironmentVariables["XLL_IDE_HOST"] = hostname;
            info.EnvironmentVariables["XLL_IDE_TOKEN"] = token;
            info.EnvironmentVariables["THEIA_DEFAULT_PLUGINS"] = "local-dir:./plugins";
            info.EnvironmentVariables["THEIA_CONFIG_DIR"] = Path.Combine(Runtime.Root, "IDE", "config");
            info.EnvironmentVariables["THEIA_PARENT_PID"] = pid.ToString();
            backend = new Process { StartInfo = info, EnableRaisingEvents = true };
            backend.OutputDataReceived += (s, e) => {
                const string prefix = "XLL_IDE_PORT=";
                if (e.Data == null) return;
                if (e.Data.StartsWith(prefix) && int.TryParse(e.Data.Substring(prefix.Length), out int port))
                    listening.TrySetResult(port);
                else Runtime.Log(e.Data);
            };
            backend.ErrorDataReceived += (s, e) => { if (e.Data != null) Runtime.Log(e.Data); };
            backend.Exited += (s, e) => listening.TrySetException(new IOException("The IDE helper exited. See python-excel.log."));
            job = new ProcessJob();
            backend.Start();
            try { job.Add(backend); } catch { backend.Kill(); throw; }
            backend.BeginOutputReadLine();
            backend.BeginErrorReadLine();
            if (await Task.WhenAny(listening.Task, Task.Delay(60000)) != listening.Task)
                throw new IOException("Timed out opening the IDE. See python-excel.log.");
            int boundPort = await listening.Task;
            if (closed) return;
            origin = new Uri("http://" + hostname + ":" + boundPort + "/");
            var core = web.CoreWebView2;
            var cookie = core.CookieManager.CreateCookie("xll_" + boundPort, token, origin.Host, "/");
            cookie.IsHttpOnly = true;
            cookie.IsSecure = false;
            cookie.SameSite = CoreWebView2CookieSameSiteKind.Strict;
            core.CookieManager.AddOrUpdateCookie(cookie);
            core.ProcessFailed += (s,e) => {
                if (e.ProcessFailedKind == CoreWebView2ProcessFailedKind.BrowserProcessExited ||
                    e.ProcessFailedKind == CoreWebView2ProcessFailedKind.RenderProcessExited) {
                    ready = false;
                    closing = false;
                }
            };
            core.NavigationStarting += (s,e) => { if (!IsLocal(e.Uri)) e.Cancel = true; };
            core.NewWindowRequested += (s,e) => e.Handled = true;
            core.WebMessageReceived += Receive;
            core.Navigate(origin.AbsoluteUri);
        }

        bool IsLocal(string uri) => Uri.TryCreate(uri, UriKind.Absolute, out var target)
            && target.GetLeftPart(UriPartial.Authority) == origin.GetLeftPart(UriPartial.Authority);

        void WorkspaceFolderChanged(string path, bool added)
        {
            path = Path.GetFullPath(path);
            var settings = Runtime.LoadSettings();
            bool worksheet = settings.folders.Any(folder => Runtime.SameFolder(folder, path));
            bool package = settings.packages.Any(folder => Runtime.SameFolder(folder, path));
            void Confirm(string question, string kind)
            {
                if (MessageBox.Show(this, question + "\n\n" + path, "Excel startup folders",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                    Runtime.ChangeStartupFolder(path, kind, added);
            }
            if (added && !worksheet && !package)
                Confirm("Also load worksheet functions from this folder when Excel starts?", "folders");
            if (!added && worksheet)
                Confirm("Also remove this worksheet folder from Excel startup?", "folders");
            if (!added && package)
                Confirm("Also remove this folder from Excel's package links?", "packages");
        }

        static bool LoadsFile(Settings settings, string file) =>
            file.EndsWith(".py", StringComparison.OrdinalIgnoreCase) && settings.folders.Any(folder =>
                file.StartsWith(Path.GetFullPath(folder).TrimEnd('\\', '/') + Path.DirectorySeparatorChar,
                                StringComparison.OrdinalIgnoreCase));

        async Task RefreshSaved(string path)
        {
            string file = Path.GetFullPath(path);
            if (!LoadsFile(Runtime.LoadSettings(), file)) return;
            int version = Interlocked.Increment(ref refreshVersion);
            await Task.Delay(300).ConfigureAwait(false);
            if (version != Volatile.Read(ref refreshVersion)) return;
            await ExcelAsyncUtil.QueueAsMacroTask(() => {
                var settings = Runtime.LoadSettings();
                if (version == Volatile.Read(ref refreshVersion) && LoadsFile(settings, file))
                    Functions.Load(settings, true);
            }).ConfigureAwait(false);
        }

        async void Receive(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            if (closed || !IsLocal(e.Source) || !IsLocal(web.CoreWebView2.Source)) return;
            int id = 0;
            try {
                var req = Runtime.Json.Deserialize<IDictionary<string, object>>(e.WebMessageAsJson);
                string operation = (string)req["operation"];
                if (operation == "ready") { ready = true; return; }
                if (operation == "closeResult") {
                    if (closing && Equals(req["allow"], true)) { allowClose = true; BeginInvoke(new Action(Close)); }
                    closing = false;
                    return;
                }
                id = (int)req["id"];
                var args = (IDictionary)req["args"];
                object value = null;
                switch (operation) {
                    case "environment":
                        value = new { pid, version, python = Runtime.Python, support = Runtime.Support };
                        break;
                    case "run":
                        string path = Path.GetFullPath((string)args["path"]);
                        value = await Task.Run(() => ExcelObjects.RunScript(path));
                        break;
                    case "connect":
                        int port = Convert.ToInt32(args["port"]);
                        string access = (string)args["token"];
                        await Task.Run(() => Runtime.Request(new { op = "ideConnect", port, token = access, session = token }));
                        break;
                    case "workspaceFolder":
                        WorkspaceFolderChanged((string)args["path"], (bool)args["added"]);
                        break;
                    case "refreshSaved":
                        await RefreshSaved((string)args["path"]);
                        break;
                    case "refresh":
                        Interlocked.Increment(ref refreshVersion);
                        await ExcelAsyncUtil.QueueAsMacroTask(() => Functions.Load(Runtime.LoadSettings(), true));
                        break;
                    default: throw new ArgumentException("Unknown Excel IDE command.");
                }
                Reply(new { id, value });
            } catch (Exception err) { Reply(new { id, error = err.Message }); }
        }
        void Reply(object reply)
        {
            if (!closed && IsLocal(web.CoreWebView2.Source))
                web.CoreWebView2.PostWebMessageAsJson(Runtime.Json.Serialize(reply));
        }
    }
}
