using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace PythonExcel
{
    // WebView2 needs an STA message loop that keeps running while Excel is busy.
    internal sealed class WebWindowHost
    {
        static readonly Lazy<bool> loader = new Lazy<bool>(() => {
            CoreWebView2Environment.SetLoaderDllFolderPath(Runtime.Support);
            return true;
        });
        readonly object lifetime = new object();
        readonly Func<Form> create;
        readonly string name;
        Form window;
        bool running, closing, reopen;

        internal WebWindowHost(string name, Func<Form> create)
        {
            this.name = name;
            this.create = create;
        }

        internal static void Configure(WebView2 web, string profile)
        {
            _ = loader.Value;
            web.CreationProperties = new CoreWebView2CreationProperties {
                UserDataFolder = Path.Combine(Runtime.Root, profile)
            };
        }

        internal void Open()
        {
            Runtime.Ensure();
            lock (lifetime)
            {
                if (running)
                {
                    var form = window;
                    if (closing) reopen = true;
                    else if (form != null)
                        form.BeginInvoke(new Action(() => {
                            if (form.IsDisposed) return;
                            if (form.WindowState == FormWindowState.Minimized) form.WindowState = FormWindowState.Normal;
                            form.Activate();
                        }));
                    return;
                }
                StartWindow();
            }
        }

        void StartWindow()
        {
            closing = reopen = false;
            running = true;
            var thread = new Thread(Run) { IsBackground = true, Name = name };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        void Run()
        {
            try
            {
                using (var form = create())
                {
                    lock (lifetime)
                    {
                        if (closing) return;
                        _ = form.Handle;
                        window = form;
                    }
                    form.FormClosing += (s, e) => { if (!e.Cancel) lock (lifetime) closing = true; };
                    form.FormClosed += (s, e) => { lock (lifetime) window = null; };
                    Application.Run(form);
                }
            }
            catch (Exception error) { Runtime.Log(error.ToString()); }
            finally
            {
                lock (lifetime)
                {
                    window = null;
                    running = false;
                    if (reopen) StartWindow();
                }
            }
        }

        internal void Shutdown()
        {
            lock (lifetime)
            {
                closing = true;
                reopen = false;
                if (window != null) window.BeginInvoke(new Action(window.Close));
            }
        }
    }
}
