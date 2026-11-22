using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace PythonExcel
{
    internal sealed class PipWindow : Form
    {
        const string Host = "python-packages.example";
        const string Page = "https://" + Host + "/index.html";
        static readonly WebWindowHost host = new WebWindowHost("Manage Packages", () => new PipWindow());
        readonly WebView2 web = new WebView2 { Dock = DockStyle.Fill };
        readonly Timer updates = new Timer { Interval = 150 };
        int revision = -1;

        internal static void Open() => host.Open();
        internal static void Shutdown() => host.Shutdown();
        internal static Task<bool> CloseForUpdate(System.Threading.CancellationToken token) => host.CloseForUpdate(token);

        PipWindow()
        {
            Text = "Manage Packages";
            ClientSize = new Size(780, 590);
            MinimumSize = new Size(440, 340);
            Controls.Add(web);
            FormClosed += (s,e) => updates.Dispose();
            updates.Tick += (s,e) => Publish();
            Shown += async (s,e) => {
                try
                {
                    WebWindowHost.Configure(web, "PackagesWebView2");
                    await web.EnsureCoreWebView2Async();
                    if (IsDisposed) return;
                    var core = web.CoreWebView2;
                    core.Settings.AreHostObjectsAllowed = false;
                    web.AllowExternalDrop = false;
                    core.SetVirtualHostNameToFolderMapping(Host, Path.Combine(Runtime.Support, "packages"),
                        CoreWebView2HostResourceAccessKind.Deny);
                    core.NavigationStarting += (sender,args) => args.Cancel = args.Uri != Page;
                    core.NewWindowRequested += (sender,args) => args.Handled = true;
                    core.WebMessageReceived += Receive;
                    core.NavigationCompleted += (sender,args) => {
                        if (!args.IsSuccess && !IsDisposed) ShowError("Could not load Manage Packages. " + args.WebErrorStatus);
                    };
                    core.Navigate(Page);
                }
                catch (Exception error) { if (!IsDisposed) ShowError(error.Message); }
            };
        }

        void ShowError(string message)
        {
            updates.Stop();
            web.Hide();
            Controls.Add(new Label { Dock = DockStyle.Fill, Text = message, Padding = new Padding(14) });
        }

        void Publish()
        {
            if (web.CoreWebView2.Source != Page) return;
            string json = Pip.Changes(ref revision);
            if (json != null) web.CoreWebView2.PostWebMessageAsJson(json);
        }

        void Receive(object sender, CoreWebView2WebMessageReceivedEventArgs args)
        {
            if (args.Source != Page || web.CoreWebView2.Source != Page) return;
            try
            {
                var message = Runtime.Json.Deserialize<Dictionary<string,string>>(args.WebMessageAsJson);
                switch (message["action"])
                {
                    case "ready":
                        revision = -1;
                        Pip.Refresh();
                        updates.Start();
                        break;
                    case "refresh": Pip.Refresh(); break;
                    case "add": Pip.Submit(null, message["name"]); break;
                    case "run": Pip.Submit(message["command"], null); break;
                }
                Publish();
            }
            catch (Exception error)
            {
                web.CoreWebView2.PostWebMessageAsJson(Runtime.Json.Serialize(new { error = error.Message }));
            }
        }
    }
}
