using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PythonExcel
{
    // Each ribbon click owns a designer, including the time spent preparing its runtime.
    internal sealed class CompilerWindow
    {
        static readonly HashSet<CompilerWindow> open = new HashSet<CompilerWindow>();
        Process process;
        EventWaitHandle withdrawn;
        TaskCompletionSource<bool> close;

        internal static void Open(string template)
        {
            var designer = new CompilerWindow();
            lock (open) open.Add(designer);
            Task.Run(() => designer.Run(template));
        }

        async Task Run(string template)
        {
            try
            {
                Runtime.Install();
                string eventName = @"Local\XllPython.CompilerClose." + Guid.NewGuid().ToString("N");
                var info = new ProcessStartInfo(Path.Combine(Runtime.Support, "Designer.exe"),
                    Published.Quote(Runtime.Python) + " " + Published.Quote(template) + " " +
                    Published.Quote(Path.Combine(Runtime.Root, "CompilerWebView2")) + " " + eventName) {
                    UseShellExecute = false, WorkingDirectory = Runtime.Support,
                    RedirectStandardInput = true, RedirectStandardOutput = true, CreateNoWindow = true
                };
                using (var child = new Process { StartInfo = info })
                {
                    lock (open)
                    {
                        if (close != null) return;
                        withdrawn = new EventWaitHandle(false, EventResetMode.ManualReset, eventName);
                        child.Start();
                        process = child;
                    }
                    string reply;
                    while ((reply = await child.StandardOutput.ReadLineAsync().ConfigureAwait(false)) != null)
                        if (reply == "CANCELED") lock (open)
                        {
                            close?.TrySetResult(false);
                            close = null;
                        }
                    child.WaitForExit();
                    lock (open) process = null;
                }
            }
            catch (Exception e) { Runtime.Log(e.ToString()); MessageBox.Show(e.Message, "Compile to XLL"); }
            finally
            {
                lock (open)
                {
                    open.Remove(this);
                    withdrawn?.Dispose();
                    close?.TrySetResult(true);
                }
            }
        }

        internal static CompilerWindow[] Snapshot() { lock (open) return open.ToArray(); }

        internal async Task<bool> CloseForUpdate(CancellationToken token)
        {
            Task<bool> result;
            lock (open)
            {
                if (!open.Contains(this)) return true;
                if (close != null) throw new InvalidOperationException("Finish the open compiler close prompt before retrying.");
                close = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
                result = close.Task;
                withdrawn?.Reset();
                // A closing designer may already have disconnected. Run confirms its exit.
                try { process?.StandardInput.WriteLine("CLOSE"); }
                catch (IOException) { }
            }
            using (token.Register(() => { lock (open) if (open.Contains(this) && close?.Task == result) withdrawn?.Set(); }))
                return await result.ConfigureAwait(false);
        }
    }
}
