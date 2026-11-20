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
                var discovered = (IDictionary)Runtime.Request(new { op = "discover", settings.folders, settings.packages });
                foreach (IDictionary function in (IList)discovered["functions"])
                    Functions.Register(function);
                foreach (string error in (IList)discovered["errors"]) Runtime.Log(error);
                Runtime.Log("Loaded " + ((IList)discovered["functions"]).Count + " Python worksheet functions.");
                if (((IList)discovered["errors"]).Count != 0)
                    ExcelDna.Logging.LogDisplay.WriteLine("Python startup errors. See " + Runtime.Root + @"\python-excel.log");
            }
            catch (Exception e)
            {
                Runtime.Log(e.ToString());
                ExcelDna.Logging.LogDisplay.WriteLine("Python startup failed: " + e.Message);
            }
        }

        public void AutoClose() { IdeWindow.Shutdown(); }
    }

    internal static class Functions
    {
        internal static void Register(IDictionary f)
        {
            string name = (string)f["name"];
            bool async = (bool)f["asynchronous"];
            var inputs = ((IList)f["arguments"]).Cast<IDictionary>().ToList();
            var parameters = inputs.Select(a => Expression.Parameter(typeof(object), (string)a["name"])).ToList();
            var callArgs = new List<Expression> { Expression.Constant(name), Expression.NewArrayInit(typeof(object), parameters) };
            var attrs = inputs.Select(a => (object)new ExcelArgumentAttribute
                {Name=(string)a["name"], Description=(string)a["description"], AllowReference=(bool)a["reference"]}).ToList();
            if (async)
            {
                var handle = Expression.Parameter(typeof(ExcelAsyncHandle), "asyncHandle");
                parameters.Add(handle);
                callArgs.Add(handle);
                attrs.Add(new ExcelArgumentAttribute());
            }
            var call = Expression.Call(typeof(Functions), async ? nameof(CallAsync) : nameof(Call), null, callArgs.ToArray());
            var excel = new ExcelFunctionAttribute {
                Name = name, Description = (string)f["description"], Category = (string)f["category"],
                IsThreadSafe = (bool)f["thread_safe"], IsMacroType = (bool)f["macro_type"],
                IsHidden = (bool)f["hidden"], IsVolatile = (bool)f["volatile"], HelpTopic = (string)f["help_topic"],
                IsClusterSafe = (bool)f["cluster_safe"]
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

        public static void CallAsync(string name, object[] arguments, ExcelAsyncHandle handle)
        {
            Task.Run(() =>
            {
                try { handle.SetResult(Call(name, arguments)); }
                catch (Exception e) { Runtime.Log(name + "\n" + e.Message); }
            });
        }
    }
}

