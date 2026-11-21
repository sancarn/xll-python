using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Security.Authentication;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using PythonExcel.Update;

namespace PythonExcel
{
    internal static class Updates
    {
        const string Repository = "sancarn/xll-python";
        const string AssetName = "PythonExcel.xll";
        const string Caption = "Python for Excel update";
        static int busy;

        sealed class Asset
        {
            public string name { get; set; }
            public string digest { get; set; }
            public long size { get; set; }
            public string browser_download_url { get; set; }
        }

        sealed class Release
        {
            public string tag_name { get; set; }
            public Asset[] assets { get; set; }
        }

        static Version VersionOf(string tag)
        {
            if (!Regex.IsMatch(tag ?? "", @"^v\d+\.\d+\.\d+$"))
                throw new InvalidDataException("Unrecognized release version: " + tag);
            return Version.Parse(tag.Substring(1));
        }

        internal static async void Check()
        {
            if (Interlocked.Exchange(ref busy, 1) != 0) return;
            string staged = null;
            UpdateProgress progress = null;
            void ShowResult(Action showDialog)
            {
                ExcelAsyncUtil.QueueAsMacro(() => {
                    progress?.Dispose();
                    showDialog();
                });
            }

            try
            {
                progress = new UpdateProgress();
                progress.Show(new Owner());
                progress.Refresh();
                string target = ExcelDnaUtil.XllPath;
                string version;
                using (var reader = new StreamReader(typeof(Updates).Assembly.GetManifestResourceStream("release.txt")))
                    version = reader.ReadToEnd().Trim();
                using (var client = new HttpClient(new HttpClientHandler { SslProtocols = SslProtocols.Tls12 })
                    { Timeout = TimeSpan.FromMinutes(3) })
                {
                    client.DefaultRequestHeaders.UserAgent.ParseAdd("PythonExcel-Updater/1.0");
                    var json = await client.GetStringAsync("https://api.github.com/repos/" + Repository + "/releases/latest").ConfigureAwait(false);
                    var release = Runtime.Json.Deserialize<Release>(json);
                    var asset = release.assets.Single(a => a.name == AssetName);
                    if (!Regex.IsMatch(asset.digest ?? "", "^sha256:[a-f0-9]{64}$") || asset.size <= 0)
                        throw new InvalidDataException("GitHub did not provide a valid SHA-256 digest and size.");
                    string hash = asset.digest.Substring(7);
                    string original = RestartPlan.Hash(target);
                    if (VersionOf(release.tag_name) < VersionOf(version) || original == hash)
                    {
                        ShowResult(() => Report("You already have the latest available update."));
                        return;
                    }
                    string url = "https://github.com/" + Repository + "/releases/download/" + release.tag_name + "/" + AssetName;
                    if (asset.browser_download_url != url) throw new InvalidDataException("Unexpected GitHub asset URL.");
                    staged = target + "." + Guid.NewGuid().ToString("N") + ".tmp";
                    using (var response = await client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(false))
                    {
                        response.EnsureSuccessStatusCode();
                        using (var output = new FileStream(staged, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                        using (var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(3)))
                        using (var source = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
                            await source.CopyToAsync(output, 81920, deadline.Token).ConfigureAwait(false);
                    }
                    if (new FileInfo(staged).Length != asset.size || RestartPlan.Hash(staged) != hash)
                        throw new InvalidDataException("The downloaded add-in failed verification.");

                    var plan = new RestartPlan { Target=target, Staged=staged, OriginalHash=original, ReleaseHash=hash };
                    ShowResult(() => Confirm(plan, release.tag_name));
                    staged = null; // Confirm or the helper owns the verified file.
                }
            }
            catch (Exception e)
            {
                Discard(staged);
                ShowResult(() => Report("Could not check or download the update.\n" + e.GetBaseException().Message));
            }
        }

        static void Discard(string path)
        {
            if (path == null) return;
            try { File.Delete(path); }
            catch (IOException e) { Runtime.Log(e.Message); }
            catch (UnauthorizedAccessException e) { Runtime.Log(e.Message); }
        }

        static void Report(string message)
        {
            try { MessageBox.Show(new Owner(), message, Caption); }
            finally { Interlocked.Exchange(ref busy, 0); }
        }

        static void Confirm(RestartPlan plan, string version)
        {
            Process helper = null;
            bool handedOff = false;
            try
            {
                var choice = ChooseWorkbookAction(version);
                if (choice != DialogResult.Yes && choice != DialogResult.No) return;

                dynamic app = ExcelDnaUtil.Application;
                plan.Files = PrepareWorkbooks(app, plan.Target, choice == DialogResult.Yes);
                plan.Arguments(); // Validate the relaunch before closing anything.
                using (var current = Process.GetCurrentProcess()) plan.ParentPid = current.Id;
                helper = Process.Start(new ProcessStartInfo(HelperPath()) {
                    UseShellExecute=false, CreateNoWindow=true, RedirectStandardInput=true, RedirectStandardOutput=true,
                    StandardOutputEncoding=Encoding.UTF8
                });
                // ASCII JSON avoids the system console code page changing Unicode workbook paths.
                string json = Regex.Replace(Runtime.Json.Serialize(plan), @"[^\u0000-\u007F]",
                    m => "\\u" + ((int)m.Value[0]).ToString("x4"));
                helper.StandardInput.WriteLine(json);
                helper.StandardInput.Close();
                string ack = helper.StandardOutput.ReadLine();
                if (ack != "READY") throw new IOException(ack ?? "The restart helper could not start.");
                var child = helper;
                child.Exited += (s,e) => Interlocked.Exchange(ref busy, 0);
                child.EnableRaisingEvents = true;
                if (choice == DialogResult.No)
                {
                    foreach (dynamic book in OpenWorkbooks(app))
                    {
                        int count = app.Workbooks.Count;
                        book.Close(SaveChanges: false);
                        if (app.Workbooks.Count >= count)
                            throw new IOException("A workbook stayed open. The update was canceled.");
                    }
                }
                handedOff = true;
                try { app.Quit(); }
                catch (COMException e) when ((uint)e.ErrorCode == 0x80010108) { return; }
                catch { handedOff = false; throw; }
                // Quit is deferred until the ribbon macro returns. Check for a workbook veto
                // on a later idle turn, after Excel has processed its pending shutdown.
                Task.Delay(2000).ContinueWith(_ => {
                    if (!child.HasExited) ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        dynamic live = ExcelDnaUtil.Application;
                        if (live.Workbooks.Count > 0 && !child.HasExited)
                        {
                            child.Kill();
                            child.WaitForExit();
                            Discard(plan.Staged);
                        }
                    });
                });
            }
            catch (OperationCanceledException) { }
            catch (Exception e) { MessageBox.Show(new Owner(), e.Message, Caption); }
            finally
            {
                if (!handedOff)
                {
                    if (helper != null)
                    {
                        if (!helper.HasExited) { helper.Kill(); helper.WaitForExit(); }
                        helper.Dispose();
                    }
                    Discard(plan.Staged);
                    Interlocked.Exchange(ref busy, 0);
                }
            }
        }

        static DialogResult ChooseWorkbookAction(string version)
        {
            using (var dialog = new Form {
                Text = Caption, FormBorderStyle = FormBorderStyle.FixedDialog,
                AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                StartPosition = FormStartPosition.CenterParent, MinimizeBox = false, MaximizeBox = false,
                ShowInTaskbar = false
            })
            {
                var layout = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill,
                    FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(12) };
                layout.Controls.Add(new Label { AutoSize = true, Text = "Excel will restart to install " + version +
                    ".\nSave changes to your open workbooks before restarting?", Margin = new Padding(3, 3, 3, 12) });
                var buttons = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
                buttons.Controls.Add(new Button { Text = "Save all", AutoSize = true, DialogResult = DialogResult.Yes });
                buttons.Controls.Add(new Button { Text = "Don't save any", AutoSize = true, DialogResult = DialogResult.No });
                var cancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
                buttons.Controls.Add(cancel);
                layout.Controls.Add(buttons);
                dialog.Controls.Add(layout);
                dialog.CancelButton = cancel;
                dialog.ActiveControl = cancel;
                return dialog.ShowDialog(new Owner());
            }
        }

        static List<object> OpenWorkbooks(dynamic app)
        {
            var books = new List<object>();
            foreach (object book in app.Workbooks) books.Add(book);
            return books;
        }

        static string[] PrepareWorkbooks(dynamic app, string target, bool save)
        {
            if (app.ProtectedViewWindows.Count != 0)
                throw new IOException("Close the Protected View windows before restarting so you can choose how to reopen them.");
            if (save)
            {
                foreach (dynamic book in OpenWorkbooks(app))
                {
                    if (String.IsNullOrEmpty(book.Path) || (!book.Saved && book.ReadOnly))
                    {
                        book.Activate();
                        if (!(bool)app.Dialogs[5].Show()) throw new OperationCanceledException(); // xlDialogSaveAs
                    }
                    else if (!book.Saved) book.Save();
                }
            }
            var files = new List<string>();
            bool installed = false;
            foreach (dynamic addin in app.AddIns)
                if (addin.Installed && String.Equals(addin.FullName, target, StringComparison.OrdinalIgnoreCase)) installed = true;
            if (!installed) files.Add(target);
            string[] automatic = { app.StartupPath, app.AltStartupPath, Path.Combine((string)app.Path, "XLSTART") };
            foreach (dynamic book in app.Workbooks)
            {
                if (save && (!book.Saved || String.IsNullOrEmpty(book.Path)))
                    throw new IOException("Excel could not save " + book.Name + ". The update was canceled.");
                if (!String.IsNullOrEmpty(book.Path) && !automatic.Contains((string)book.Path, StringComparer.OrdinalIgnoreCase))
                    files.Add((string)book.FullName);
            }
            return files.ToArray();
        }

        static string HelperPath()
        {
            string dir = Path.Combine(Runtime.Root, "updaters");
            Directory.CreateDirectory(dir);
            string temp = Path.Combine(dir, Guid.NewGuid().ToString("N") + ".tmp");
            try
            {
                using (var resource = typeof(Updates).Assembly.GetManifestResourceStream("updater.exe"))
                using (var file = File.Create(temp)) resource.CopyTo(file);
                string hash = RestartPlan.Hash(temp);
                string path = Path.Combine(dir, hash + ".exe");
                try { File.Move(temp, path); }
                catch (IOException) when (File.Exists(path) && RestartPlan.Hash(path) == hash) { }
                return path;
            }
            finally { File.Delete(temp); }
        }
    }
}
