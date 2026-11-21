using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace PythonExcel
{
    // One command or console owns the pipe and its Excel object references.
    internal sealed class ExcelObjects : IDisposable
    {
        internal readonly string Name = "XllPython-" + Guid.NewGuid().ToString("N");
        readonly NamedPipeServerStream pipe;
        readonly Dictionary<int, object> objects = new Dictionary<int, object>();
        int next;
        volatile bool closed;

        internal ExcelObjects()
        {
            var security = new PipeSecurity();
            security.AddAccessRule(new PipeAccessRule(WindowsIdentity.GetCurrent().User,
                PipeAccessRights.FullControl, AccessControlType.Allow));
            pipe = new NamedPipeServerStream(Name, PipeDirection.InOut, 1, PipeTransmissionMode.Byte,
                PipeOptions.Asynchronous, 0, 0, security);
            _ = Serve();
        }

        internal static string RunScript(string path)
        {
            using (var session = new ExcelObjects())
                return (string)Runtime.Request(new { op="script", path, excelPipe=session.Name });
        }

        async Task Serve()
        {
            try
            {
                await pipe.WaitForConnectionAsync().ConfigureAwait(false);
                using (var input = new StreamReader(pipe, Encoding.UTF8, false, 4096, true))
                using (var output = new StreamWriter(pipe, new UTF8Encoding(false), 4096, true) { AutoFlush=true })
                {
                    string line;
                    while ((line = await input.ReadLineAsync().ConfigureAwait(false)) != null)
                    {
                        Dictionary<string,object> request;
                        try { request = Runtime.Json.Deserialize<Dictionary<string,object>>(line); }
                        catch (Exception e) {
                            await output.WriteLineAsync(Runtime.Json.Serialize(new { error=e.Message })).ConfigureAwait(false);
                            continue;
                        }
                        string reply = null;
                        // Perform COM work only in Excel's macro context. Serialize there too.
                        await ExcelAsyncUtil.QueueAsMacroTask(() => {
                            if (closed) return;
                            try { reply = Runtime.Json.Serialize(new { value=Invoke(request) }); }
                            catch (Exception e) { reply = Runtime.Json.Serialize(new { error=e.GetBaseException().Message }); }
                        }).ConfigureAwait(false);
                        if (closed) return;
                        await output.WriteLineAsync(reply).ConfigureAwait(false);
                    }
                }
            }
            catch (Exception e) { if (!closed) Runtime.Log("Excel object session: " + e.Message); }
            finally { pipe.Dispose(); objects.Clear(); }
        }

        object Argument(object value)
        {
            if (value is IDictionary reference) return objects[Convert.ToInt32(reference["handle"])];
            if (!(value is IList rows)) return value;
            if (rows.Count == 0 || !(rows[0] is IList first))
                throw new ArgumentException("Range values must be a rectangular list of rows.");
            var matrix = new object[rows.Count, first.Count];
            for (int r=0; r<rows.Count; r++)
            {
                var row = rows[r] as IList;
                if (row == null || row.Count != first.Count) throw new ArgumentException("Unequal row lengths.");
                for (int c=0; c<row.Count; c++) matrix[r,c] = row[c];
            }
            return matrix;
        }

        object Invoke(Dictionary<string,object> request)
        {
            foreach (var id in (IList)request["release"]) objects.Remove(Convert.ToInt32(id));
            int target = Convert.ToInt32(request["target"]);
            object instance = target == 0 ? ExcelDnaUtil.Application : objects[target];
            string action = (string)request["action"];
            var flags = action == "get" ? BindingFlags.GetProperty : action == "set"
                ? BindingFlags.SetProperty : BindingFlags.InvokeMethod;
            var args = ((IList)request["args"]).Cast<object>().Select(Argument).ToArray();
            var value = instance.GetType().InvokeMember((string)request["member"], flags, null, instance,
                args, CultureInfo.GetCultureInfo("en-US"));
            if (value != null && Marshal.IsComObject(value))
            {
                objects.Add(++next, value);
                return new { handle=next };
            }
            if (value is Array matrix && matrix.Rank == 2)
            {
                var rows = new object[matrix.GetLength(0)][];
                for (int r=0; r<rows.Length; r++)
                {
                    rows[r] = new object[matrix.GetLength(1)];
                    for (int c=0; c<rows[r].Length; c++)
                        rows[r][c] = matrix.GetValue(r+matrix.GetLowerBound(0), c+matrix.GetLowerBound(1));
                }
                return rows;
            }
            return value;
        }

        public void Dispose()
        {
            closed = true;
            pipe.Dispose();
        }
    }
}
