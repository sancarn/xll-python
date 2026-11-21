using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace PythonExcel
{
    internal sealed class Published : IDisposable
    {
        internal static Published Current;
        internal readonly IDictionary Manifest;
        readonly object gate = new object();
        Worker worker;

        Published(IDictionary info) { Manifest = info; }

        internal static bool Open()
        {
            var data = Runtime.Resource("PYTHON_MANIFEST", true);
            if (data == null) return false;
            Current = new Published((IDictionary)Runtime.Json.DeserializeObject(Encoding.UTF8.GetString(data)));
            Functions.Synchronize((IList)Current.Manifest["functions"], false);
            return true;
        }

        internal object Image(string id)
        {
            var data = (string)((IDictionary)Manifest["images"])[id];
            if (data == null) return null;
            using (var stream = new MemoryStream(Convert.FromBase64String(data)))
            using (var source = new Bitmap(stream)) return new Bitmap(source);
        }

        internal object Request(object message)
        {
            Worker instance;
            lock (gate)
            {
                if (worker == null)
                {
                    Runtime.Install(false);
                    string path = Runtime.Unpack("project", (string)Manifest["hash"]);
                    worker = new Worker(path);
                }
                instance = worker;
            }
            return instance.Request(message);
        }

        internal async void Command(string id)
        {
            try
            {
                var target = (string)((IDictionary)Manifest["actions"])[id];
                using (var objects = new ExcelObjects())
                    await Task.Run(() => Request(new {op="command", name=target, excelPipe=objects.Name}));
            }
            catch (Exception e)
            {
                Runtime.Log(e.ToString());
                System.Windows.Forms.MessageBox.Show(e.Message, (string)Manifest["name"]);
            }
        }

        public void Dispose()
        {
            lock (gate) { worker?.Dispose(); worker = null; }
        }

        sealed class Worker : IDisposable
        {
            readonly Process process;
            readonly object gate = new object();
            readonly Dictionary<long,TaskCompletionSource<string>> pending = new Dictionary<long,TaskCompletionSource<string>>();
            long sequence;
            Exception failure;

            internal Worker(string path)
            {
                var info = Runtime.PipProcess("-I -S -B -u " + Quote(Path.Combine(Runtime.Support, "_xll_worker.py"))
                    + " serve " + Quote(path));
                info.WorkingDirectory = Path.Combine(path, "src");
                info.RedirectStandardInput = true;
                process = Process.Start(info);
                process.ErrorDataReceived += (s,e) => { if (!String.IsNullOrEmpty(e.Data)) Runtime.Log(e.Data); };
                process.BeginErrorReadLine();
                _ = ReadReplies();
            }

            async Task ReadReplies()
            {
                try
                {
                    string line;
                    while ((line = await process.StandardOutput.ReadLineAsync().ConfigureAwait(false)) != null)
                    {
                        var reply = Runtime.Json.Deserialize<Dictionary<string,object>>(line);
                        long id = Convert.ToInt64(reply["id"]);
                        lock (gate)
                        {
                            var waiter = pending[id];
                            pending.Remove(id);
                            waiter.TrySetResult(line);
                        }
                    }
                    throw new IOException("The add-in's Python worker exited. Reload the add-in to restart it. See python-excel.log.");
                }
                catch (Exception e)
                {
                    lock (gate)
                    {
                        failure = e;
                        foreach (var waiter in pending.Values) waiter.TrySetException(e);
                        pending.Clear();
                    }
                }
            }

            internal object Request(object request)
            {
                var waiter = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
                lock (gate)
                {
                    if (failure != null) throw new IOException(failure.Message, failure);
                    long id = ++sequence;
                    string message = Runtime.Json.Serialize(new { id, request });
                    pending.Add(id, waiter);
                    try {
                        byte[] bytes = Encoding.UTF8.GetBytes(message + "\n");
                        process.StandardInput.BaseStream.Write(bytes, 0, bytes.Length);
                        process.StandardInput.BaseStream.Flush();
                    }
                    catch { pending.Remove(id); throw; }
                }
                var reply = Runtime.Json.Deserialize<Dictionary<string,object>>(waiter.Task.GetAwaiter().GetResult());
                if (reply.ContainsKey("failure")) throw new InvalidOperationException((string)reply["failure"]);
                return reply["value"];
            }

            public void Dispose()
            {
                lock (gate)
                {
                    try { process.StandardInput.Close(); }
                    finally
                    {
                        if (!process.HasExited) process.Kill();
                        process.Dispose();
                    }
                }
            }
        }

        internal static string Quote(string path) => "\"" + path + "\"";
    }
}
