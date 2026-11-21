using System;
using System.Collections;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.Integration.Extensibility;

namespace PythonExcel
{
#if !RUNTIME_ONLY
    [ComVisible(true), Guid("F656D923-6408-4308-8245-07401C95FBEA"), ProgId("XllPython.Automation")]
    public sealed class WorkbenchAutomation : ExcelComAddIn
    {
        public WorkbenchAutomation() { FriendlyName = "XLL Python Automation"; }

        public override void OnConnection(object app, ext_ConnectMode mode, object instance, ref Array custom)
        {
            base.OnConnection(app, mode, instance, ref custom);
            ((dynamic)instance).Object = new PythonAutomation();
        }
    }
#endif

    public partial class PythonRibbon
    {
        public override void OnConnection(object app, ext_ConnectMode mode, object instance, ref Array custom)
        {
            base.OnConnection(app, mode, instance, ref custom);
            if (Published.Current != null) ((dynamic)instance).Object = new PythonAutomation();
        }
    }

    [ComVisible(true), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IPythonAutomation
    {
        [DispId(1)] object CallFunction(string name, [Optional] object arguments);
        [DispId(2)] PythonJob StartMacro(string name);
    }

    [ComVisible(true), ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(IPythonAutomation))]
    public sealed class PythonAutomation : IPythonAutomation
    {
        public object CallFunction(string name, object arguments)
        {
            var entry = Functions.Lookup(name);
            var parameters = (IList)entry["arguments"];
            var values = arguments as Array;
            if (!(arguments is Missing) && (values == null || values.Rank != 1))
                throw new ArgumentException("Pass the positional arguments in a VBA Array(...).");
            int count = values?.Length ?? 0;
            if (count > parameters.Count) throw new ArgumentException("Too many arguments for " + name);
            var args = new object[count];
            for (int i = 0; i < count; i++)
            {
                var value = values.GetValue(i + values.GetLowerBound(0));
                args[i] = Argument(value, (bool)((IDictionary)parameters[i])["reference"]);
            }
            if ((bool)entry["generator"]) return new PythonSubscription(name, args);
            Func<object> call = () => AutomationValue(Functions.ToExcel(Runtime.Request(new { op="call", name, args })));
            if ((bool)entry["asynchronous"]) return new PythonJob(call);
            return call();
        }

        static object Argument(object value, bool reference, bool fromRange = false)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                dynamic range = value;
                if (!reference) return Argument((object)range.Value2, false, true);
                string sheet = "[" + range.Parent.Parent.Name + "]" + range.Parent.Name;
                var areas = new int[(int)range.Areas.Count][];
                for (int i = 0; i < areas.Length; i++)
                {
                    dynamic area = range.Areas[i + 1];
                    int row = area.Row - 1, col = area.Column - 1;
                    areas[i] = new[] { row, row + (int)area.Rows.Count - 1, col, col + (int)area.Columns.Count - 1 };
                }
                return new { sheet_id=new ExcelReference(0, 0, 0, 0, sheet).SheetId.ToInt64(), areas };
            }
            if (value is Array a)
            {
                if (a.Rank == 1) return a.Cast<object>().Select(v => Argument(v, false, fromRange)).ToArray();
                if (a.Rank != 2) throw new ArgumentException("Expected a vector or rectangular array.");
                var rows = new object[a.GetLength(0)][];
                for (int r = 0; r < rows.Length; r++)
                {
                    rows[r] = new object[a.GetLength(1)];
                    for (int c = 0; c < rows[r].Length; c++)
                        rows[r][c] = Argument(a.GetValue(r + a.GetLowerBound(0), c + a.GetLowerBound(1)), false, fromRange);
                }
                return rows;
            }
            if (value is DateTime date) return date.ToOADate();
            if (value is Missing) return new { missing=true };
            // Range.Value2 supplies numbers as doubles and errors as boxed HRESULTs.
            // A caller's own VBA Long can have exactly the same bits as an HRESULT.
            if (fromRange && value is int hr && (hr & 0xffff0000) == 0x800a0000)
            {
                int code = (hr & 0xffff) - 2000;
                if (Enum.IsDefined(typeof(ExcelError), (ExcelError)code)) return new { error=code };
            }
            return value is DBNull ? null : value;
        }

        internal static object AutomationValue(object value)
        {
            if (value is ExcelEmpty) return null;
            if (value is ExcelError error) return new ErrorWrapper(unchecked((int)0x800A0000) + 2000 + (int)error);
            if (value is object[,] matrix)
            {
                var result = Array.CreateInstance(typeof(object), new[] { matrix.GetLength(0), matrix.GetLength(1) }, new[] { 1, 1 });
                for (int r = 0; r < matrix.GetLength(0); r++)
                    for (int c = 0; c < matrix.GetLength(1); c++)
                        result.SetValue(AutomationValue(matrix[r,c]), r + 1, c + 1);
                return result;
            }
            return value;
        }

        public PythonJob StartMacro(string name)
        {
            var published = Published.Current;
            string target = name;
#if !RUNTIME_ONLY
            if (published == null)
            {
                var matches = Functions.Macros.Cast<IDictionary>().Where(m => {
                    string source = (string)m["source"];
                    string member = ((string)m["target"]).Split('.').Last();
                    return String.Equals(Path.GetFileNameWithoutExtension(source) + "." + member, name, StringComparison.OrdinalIgnoreCase)
                        || String.Equals(source + "!" + member, name, StringComparison.OrdinalIgnoreCase);
                }).ToArray();
                if (matches.Length != 1)
                    throw new ArgumentException((matches.Length == 0 ? "Unknown" : "Ambiguous") + " Python macro: " + name);
                target = (string)matches[0]["target"];
            }
#endif
            return new PythonJob(() => {
                using (var session = new ExcelObjects())
                {
                    if (published != null) return published.Request(new { op="macro", name=target, excelPipe=session.Name });
                    return Runtime.Request(new { op="macro", target, excelPipe=session.Name });
                }
            });
        }
    }

    [ComVisible(true), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IPythonJob
    {
        [DispId(1)] bool Done { get; }
        [DispId(2)] string Error { get; }
        [DispId(3)] object Result { get; }
    }

    [ComVisible(true), ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(IPythonJob))]
    public sealed class PythonJob : IPythonJob
    {
        readonly Task<object> task;
        internal PythonJob(Func<object> run) { task = Task.Run(run); }
        public bool Done => task.IsCompleted;
        public string Error => task.Exception?.GetBaseException().Message ?? "";
        public object Result => Done ? task.GetAwaiter().GetResult() : throw new InvalidOperationException("Python is still running.");
    }

    [ComVisible(true), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IPythonSubscription
    {
        [DispId(1)] bool HasValue { get; }
        [DispId(2)] object Value { get; }
        [DispId(3)] void Close();
    }

    [ComVisible(true), ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(IPythonSubscription))]
    public sealed class PythonSubscription : IPythonSubscription, IExcelObserver
    {
        readonly IDisposable subscription;
        volatile Tuple<object> latest;

        internal PythonSubscription(string name, object[] args)
        {
            subscription = new LiveFunction(name, args).Subscribe(this);
        }

        public bool HasValue => latest != null;
        public object Value => latest?.Item1;
        public void Close() => subscription.Dispose();

        void IExcelObserver.OnNext(object value) =>
            latest = Tuple.Create(PythonAutomation.AutomationValue(((Tuple<object>)value).Item1));
        void IExcelObserver.OnError(Exception error) { }
        void IExcelObserver.OnCompleted() { }
    }
}
