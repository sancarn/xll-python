using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace PythonExcel
{
    internal sealed class ConsoleWindow : Form
    {
        const string Host = "python-console.example";
        const string Page = "https://" + Host + "/index.html";
        static readonly WebWindowHost host = new WebWindowHost("Excel Python Console", () => new ConsoleWindow());
        readonly string session = Guid.NewGuid().ToString("N");
        readonly WebView2 web = new WebView2 { Dock = DockStyle.Fill };
        readonly SemaphoreSlim requests = new SemaphoreSlim(1);
        readonly ExcelObjects excel = new ExcelObjects();
        bool sessionOpen;
        volatile bool closed;

        internal static void Open() => host.Open();
        internal static void Shutdown() => host.Shutdown();

        ConsoleWindow()
        {
            Text = "Python Console";
            ClientSize = new Size(800, 560);
            MinimumSize = new Size(440, 320);
            Controls.Add(web);
            Shown += async (s, e) => await Initialize();
            FormClosed += (s, e) => {
                closed = true;
                excel.Dispose();
                if (sessionOpen)
                {
                    try
                    {
                        if ((bool)Call("detach"))
                            ExcelAsyncUtil.QueueAsMacro(Functions.ConsoleChanged);
                    }
                    catch (Exception error) { Runtime.Log(error.ToString()); }
                    finally { QueueRequest(null); }
                }
            };
        }

        async Task Initialize()
        {
            try
            {
                WebWindowHost.Configure(web, "ConsoleWebView2");
                await web.EnsureCoreWebView2Async();
                if (IsDisposed) return;
                var core = web.CoreWebView2;
                core.Settings.AreHostObjectsAllowed = false;
                core.Settings.AreDefaultContextMenusEnabled = false;
                core.Settings.AreBrowserAcceleratorKeysEnabled = false;
                web.AllowExternalDrop = false;
                core.SetVirtualHostNameToFolderMapping(Host, Path.Combine(Runtime.Support, "repl"),
                    CoreWebView2HostResourceAccessKind.Deny);
                core.NavigationStarting += (s, e) => e.Cancel = e.Uri != Page;
                core.NewWindowRequested += (s, e) => e.Handled = true;
                core.WebMessageReceived += MessageReceived;
                core.NavigationCompleted += (s, e) => {
                    if (!e.IsSuccess && !IsDisposed) ShowError("Could not load the console: " + e.WebErrorStatus);
                };
                Call("open");
                sessionOpen = true;
                core.Navigate(Page);
            }
            catch (Exception error) { if (!IsDisposed) ShowError(error.Message); }
        }

        void ShowError(string message)
        {
            Runtime.Log(message);
            web.Hide();
            Controls.Add(new Label { Dock = DockStyle.Fill, Padding = new Padding(12), Text = message });
        }

        object Call(string action, object args = null) =>
            Runtime.Request(new { op = "repl", session, action, args, excelPipe=excel.Name });

        void Send(object reply)
        {
            if (closed) return;
            string json = Runtime.Json.Serialize(reply);
            try
            {
                BeginInvoke(new Action(() => {
                    if (!closed && web.CoreWebView2.Source == Page) web.CoreWebView2.PostWebMessageAsJson(json);
                }));
            }
            catch (InvalidOperationException) when (closed) { }
        }

        void Drain(int id)
        {
            var events = (System.Collections.IList)Call("drain");
            if (events.Count != 0) Send(new { id, events });
        }

        void MessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            if (e.Source != Page || web.CoreWebView2.Source != Page) return;
            QueueRequest(e.WebMessageAsJson);
        }

        void QueueRequest(string json)
        {
            var turn = requests.WaitAsync();
            _ = Task.Run(async () => {
                await turn;
                try { await ProcessRequest(json); }
                finally { requests.Release(); }
            });
        }

        async Task ProcessRequest(string json)
        {
            int id = 0;
            try
            {
                if (json == null) { Call("close"); return; }
                if (closed) return;
                var message = Runtime.Json.Deserialize<Dictionary<string, object>>(json);
                id = (int)message["id"];
                string action = (string)message["action"];
                if (action != "run" && action != "complete" && action != "children" && action != "clear")
                    throw new ArgumentException("Unknown console request");
                object args = message["args"];
                object value;
                if (action == "run")
                {
                    var work = Task.Run(() => Call(action, args));
                    while (!closed && await Task.WhenAny(work, Task.Delay(80)) != work)
                        Drain(id);
                    value = await work;
                    if (!closed) Drain(id);
                    if (((System.Collections.IDictionary)value).Contains("worksheetFormulas"))
                        await ExcelAsyncUtil.QueueAsMacroTask(Functions.ConsoleChanged);
                }
                else value = Call(action, args);
                Send(new { id, value });
            }
            catch (Exception error) { Send(new { id, error = error.Message }); }
        }
    }
}
