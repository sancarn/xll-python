using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace PythonExcel
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            try
            {
                var settings = Runtime.LoadSettings();
                if (settings.folders.Length != 0 || settings.packages.Length != 0)
                    Functions.Load(settings, false);
            }
            catch (Exception e)
            {
                Runtime.Log(e.ToString());
                ExcelDna.Logging.LogDisplay.WriteLine("Python startup failed: " + e.Message);
            }
        }

        public void AutoClose() { ConsoleWindow.Shutdown(); IdeWindow.Shutdown(); }
    }

    internal static class Functions
    {
        sealed class Registration
        {
            internal string Signature;
            internal double Id;
        }

        static readonly Dictionary<string,Registration> registered = new Dictionary<string,Registration>(StringComparer.OrdinalIgnoreCase);
        static int generation;

        internal static void Load(Settings settings, bool refresh)
        {
            var discovered = (IDictionary)Runtime.Request(new {op="discover", settings.folders, settings.packages, refresh});
            var errors = (IList)discovered["errors"];
            foreach (string error in errors) Runtime.Log(error);
            if (errors.Count != 0)
            {
                ExcelDna.Logging.LogDisplay.WriteLine("Python startup errors. See " + Runtime.Root + @"\python-excel.log");
            }
            var entries = ((IList)discovered["functions"]).Cast<IDictionary>()
                .ToDictionary(f => (string)f["name"], StringComparer.OrdinalIgnoreCase);
            foreach (string name in registered.Keys.ToArray())
            {
                if (entries.TryGetValue(name, out var entry) && registered[name].Signature == Runtime.Json.Serialize(entry))
                    continue;
                // Excel's Function Wizard retains the name unless it is cleared as a command.
                XlCall.Excel(XlCall.xlfUnregister, registered[name].Id);
                var command = XlCall.Excel(XlCall.xlfRegister, ExcelDnaUtil.XllPath, "xlAutoRemove", "I", name, ExcelMissing.Value, 2);
                if (command is double)
                {
                    XlCall.Excel(XlCall.xlfSetName, name);
                    XlCall.Excel(XlCall.xlfUnregister, command);
                }
                registered.Remove(name);
            }
            foreach (var pair in entries)
                if (!registered.ContainsKey(pair.Key))
                {
                    Register(pair.Value);
                    registered.Add(pair.Key, new Registration { Signature=Runtime.Json.Serialize(pair.Value),
                        Id=(double)XlCall.Excel(XlCall.xlfEvaluate, pair.Key) });
                }
            generation++;
            Runtime.Log("Loaded " + entries.Count + " Python worksheet functions.");
            if (refresh)
            {
                dynamic app = ExcelDnaUtil.Application;
                app.CalculateFullRebuild();
            }
        }

        static void Register(IDictionary f)
        {
            string name = (string)f["name"];
            bool async = (bool)f["asynchronous"];
            var inputs = ((IList)f["arguments"]).Cast<IDictionary>().ToList();
            var parameters = inputs.Select(a => Expression.Parameter(typeof(object), (string)a["name"])).ToList();
            var callArgs = new List<Expression> { Expression.Constant(name), Expression.NewArrayInit(typeof(object), parameters) };
            var attrs = inputs.Select(a => (object)new ExcelArgumentAttribute
                {Name=(string)a["name"], Description=(string)a["description"], AllowReference=(bool)a["reference"]}).ToList();
            var call = Expression.Call(typeof(Functions), async ? nameof(CallAsync) : nameof(Call), null, callArgs.ToArray());
            var excel = new ExcelFunctionAttribute {
                Name = name, Description = (string)f["description"], Category = (string)f["category"],
                IsThreadSafe = !async && (bool)f["thread_safe"], IsMacroType = (bool)f["macro_type"],
                IsHidden = (bool)f["hidden"], IsVolatile = (bool)f["volatile"], HelpTopic = (string)f["help_topic"],
                IsClusterSafe = !async && (bool)f["cluster_safe"], SuppressOverwriteError = true
            };
            ExcelIntegration.RegisterLambdaExpressions(new List<LambdaExpression>{Expression.Lambda(call, parameters)},
                new List<object>{excel}, new List<List<object>>{attrs});
        }

        static object ToPython(object value)
        {
            if (value is ExcelMissing) return new {missing = true};
            if (value == null || value is ExcelEmpty) return null;
            if (value is ExcelError) return new {error = (int)(ExcelError)value};
            if (value is ExcelReference reference)
                return new {sheet_id=reference.SheetId.ToInt64(), areas=reference.InnerReferences.Select(
                    a => new[]{a.RowFirst,a.RowLast,a.ColumnFirst,a.ColumnLast}).ToArray()};
            var range = value as object[,];
            if (range == null) return value;
            var rows = new object[range.GetLength(0)][];
            for (int r=0; r<rows.Length; r++)
            {
                rows[r] = new object[range.GetLength(1)];
                for (int c=0; c<rows[r].Length; c++) rows[r][c] = ToPython(range[r,c]);
            }
            return rows;
        }

        static object ToExcel(object value)
        {
            if (value == null) return ExcelEmpty.Value;
            if (value is IDictionary map && map.Contains("error")) return (ExcelError)Convert.ToInt32(map["error"]);
            if (value is IList rows)
            {
                if (rows.Count == 0) return ExcelEmpty.Value;
                bool matrix = rows[0] is IList;
                int columns = matrix ? ((IList)rows[0]).Count : 1;
                if (columns == 0) return ExcelEmpty.Value;
                var array = new object[rows.Count, columns];
                for (int r=0; r<rows.Count; r++)
                {
                    if (matrix != (rows[r] is IList) || (matrix && ((IList)rows[r]).Count != columns))
                        throw new ArgumentException("Python returned a ragged array.");
                    for (int c=0; c<columns; c++)
                        array[r,c] = ToExcel(matrix ? ((IList)rows[r])[c] : rows[r]);
                }
                return array;
            }
            if (value is string || value is bool || value is double) return value;
            return Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture);
        }

        public static object Call(string name, object[] arguments)
        {
            try { return ToExcel(Runtime.Request(new { op="call", name, args=arguments.Select(ToPython).ToArray() })); }
            catch (Exception e) { Runtime.Log(name + "\n" + e.Message); return ExcelError.ExcelErrorValue; }
        }

        public static object CallAsync(string name, object[] arguments)
        {
            // Excel can finish the calculation pass while its RTD subscription is pending.
            // Wrap completed values to distinguish a legitimate Python #N/A from RTD's initial #N/A.
            var result = ExcelAsyncUtil.RunTask(name + ":" + generation, arguments,
                () => Task.Run(() => Tuple.Create(Call(name, arguments))));
            if (result is Tuple<object> completed) return completed.Item1;
            return Equals(result, ExcelError.ExcelErrorNA) ? (object)"Waiting…" : result;
        }
    }
}

