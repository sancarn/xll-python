using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace PythonExcel
{
    [ComVisible(true)]
    public class PythonRibbon : ExcelRibbon
    {
        public override string GetCustomUI(string ribbonId) => @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon><tabs><tab id='xllPythonTab' label='Python'>
    <group id='xllPythonRun' label='Python'>
      <button id='xllPythonScript' label='Run Script' size='large' imageMso='MacroPlay' onAction='RunScript'/>
      <button id='xllPythonIde' label='Open IDE' size='large' imageMso='VisualBasic' onAction='OpenIde'/>
    </group>
    <group id='xllPythonSetup' label='Environment'>
      <button id='xllPythonPip' label='Package Installer' size='large' imageMso='AddInManager' onAction='Packages'/>
      <button id='xllPythonFolders' label='Startup Folders' size='large' imageMso='FileOpen' onAction='Folders'/>
      <button id='xllPythonUpdate' label='Check for Updates' size='large' imageMso='RefreshAll' onAction='CheckUpdates'/>
    </group>
  </tab></tabs></ribbon>
</customUI>";

        static void Guard(Action action)
        {
            try { action(); }
            catch (Exception e) { Runtime.Log(e.ToString()); MessageBox.Show(e.Message, "Python for Excel"); }
        }

        public void Folders(IRibbonControl control) => Guard(() =>
        {
            using (var dialog = new FolderWindow())
                if (dialog.ShowDialog(new Owner()) == DialogResult.Retry)
                {
                    var settings = dialog.Selection;
                    ExcelAsyncUtil.QueueAsMacro(() => Guard(() => {
                        Functions.Load(settings, true);
                        Runtime.SaveSettings(settings);
                    }));
                }
        });

        public void Packages(IRibbonControl control) => Guard(() =>
        {
            Runtime.Ensure();
            new PipWindow().Show(new Owner());
        });

        public void OpenIde(IRibbonControl control) => Guard(IdeWindow.Open);
        public void CheckUpdates(IRibbonControl control) => Guard(Updates.Check);

        public void RunScript(IRibbonControl control) => Guard(() =>
        {
            using (var dialog = Ui.ScriptPicker())
                if (dialog.ShowDialog(new Owner()) == DialogResult.OK)
                {
                    var output = new OutputWindow(Path.GetFileName(dialog.FileName));
                    output.Show(new Owner());
                    output.Run(dialog.FileName);
                }
        });
    }

    internal sealed class Owner : IWin32Window
    {
        public IntPtr Handle => ExcelDnaUtil.WindowHandle;
    }

    internal static class Ui
    {
        internal static OpenFileDialog ScriptPicker() => new OpenFileDialog {
            Filter = "Python scripts (*.py;*.pyw)|*.py;*.pyw|All files (*.*)|*.*", RestoreDirectory = true };
        internal static TextBox ConsoleBox() => new TextBox { Dock=DockStyle.Fill, Multiline=true, ReadOnly=true,
            ScrollBars=ScrollBars.Both, WordWrap=false, Font=new Font(FontFamily.GenericMonospace, 10) };
        internal static Button Button(string label, EventHandler action)
        {
            var button = new Button { Text=label, AutoSize=true };
            button.Click += action;
            return button;
        }
    }

    internal sealed class OutputWindow : Form
    {
        readonly TextBox output = Ui.ConsoleBox();
        internal OutputWindow(string name)
        {
            Text="Python output — " + name; Width=740; Height=460; Controls.Add(output);
        }
        internal async void Run(string path)
        {
            output.Text="Running " + path + Environment.NewLine;
            try
            {
                string result = (string)await Task.Run(() => Runtime.Request(new {op="script", path}));
                if (!IsDisposed) output.AppendText(result.Replace("\r\n","\n").Replace("\n","\r\n") + "\r\nFinished.");
            }
            catch(Exception e) { Runtime.Log(e.Message); if (!IsDisposed) output.AppendText(e.Message); }
        }
    }
}

