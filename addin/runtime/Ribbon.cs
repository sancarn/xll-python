using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;

namespace PythonExcel
{
    [ComVisible(true)]
    public class PythonRibbon : ExcelRibbon
    {
        public override string GetCustomUI(string ribbonId) => (string)Published.Current?.Manifest["ribbon"];
        public override object LoadImage(string id) => Published.Current?.Image(id) ?? base.LoadImage(id);
        public void RunPublished(IRibbonControl control) => Published.Current.Command(control.Id);
    }
}
