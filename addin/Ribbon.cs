using System;
using System.Collections;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace PythonExcel
{
    [ComVisible(true)]
    public partial class PythonRibbon : ExcelRibbon
    {
        public override string GetCustomUI(string ribbonId) => Published.Current != null ? (string)Published.Current.Manifest["ribbon"] : @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon><tabs><tab id='xllPythonTab' label='Python'>
    <group id='xllPythonRun' label='Python'>
      <button id='xllPythonScript' label='Run Script' size='large' imageMso='PlayMacro' onAction='RunScript'/>
      <button id='xllPythonMacro' label='Run Macro' size='large' imageMso='MacroPlay' onAction='RunMacro'/>
      <button id='xllPythonCompile' label='Compile to XLL' size='large' imageMso='FileSaveAs' onAction='OpenCompiler'/>
      <button id='xllPythonConsole' label='Console' size='large' imageMso='ScriptEditor' onAction='OpenConsole'/>
      <button id='xllPythonIde' label='Open IDE' size='large' imageMso='VisualBasic' onAction='OpenIde'/>
    </group>
    <group id='xllPythonSetup' label='Environment'>
      <button id='xllPythonPip' label='Manage Packages' size='large' imageMso='AddInManager' onAction='Packages'/>
      <button id='xllPythonFolders' label='Startup Folders' size='large' imageMso='FileOpen' onAction='Folders'/>
      <button id='xllPythonUpdate' label='Check for Updates' size='large' imageMso='RefreshAll' onAction='CheckUpdates'/>
    </group>
  </tab></tabs></ribbon>
</customUI>";

        public override object LoadImage(string id) => Published.Current?.Image(id) ?? base.LoadImage(id);
        public void RunPublished(IRibbonControl control) => Published.Current.Command(control.Id);

        public void OpenCompiler(IRibbonControl control)
        {
            string template = ExcelDnaUtil.XllPath;
            CompilerWindow.Open(template);
        }

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
                    var original = dialog.Original;
                    var selection = dialog.Selection;
                    ExcelAsyncUtil.QueueAsMacro(() => Guard(() => Runtime.SaveFolderSelection(original, selection, true)));
                }
        });

        public void Packages(IRibbonControl control) => Guard(PipWindow.Open);

        public void OpenIde(IRibbonControl control) => Guard(IdeWindow.Open);
        public void OpenConsole(IRibbonControl control) => Guard(ConsoleWindow.Open);
        public void CheckUpdates(IRibbonControl control) => Guard(Updates.Check);

        public void RunScript(IRibbonControl control) => Guard(() =>
        {
            using (var dialog = Ui.ScriptPicker())
                if (dialog.ShowDialog(new Owner()) == DialogResult.OK)
                {
                    var output = new OutputWindow(Path.GetFileName(dialog.FileName));
                    output.Show(new Owner());
                    string path = dialog.FileName;
                    output.Run(path, () => ExcelObjects.RunScript(path));
                }
        });

        public void RunMacro(IRibbonControl control) => Guard(() =>
        {
            using (var picker = new Form { Text="Run Macro", Size=new Size(650, 340),
                StartPosition=FormStartPosition.CenterParent, MinimizeBox=false, MaximizeBox=false })
            {
                var list = new ListView { Dock=DockStyle.Fill, View=View.Details, FullRowSelect=true,
                    MultiSelect=false, HideSelection=false, AccessibleName="Startup macros" };
                list.Columns.Add("Macro", 200);
                list.Columns.Add("Source", 415);
                foreach (var m in Functions.Macros.Cast<IDictionary>().OrderBy(m => (string)m["label"]))
                {
                    string target = (string)m["target"];
                    list.Items.Add(new ListViewItem(new[] { (string)m["label"],
                        m["source"] + " (" + target.Substring(target.LastIndexOf('.') + 1) + ")" }) { Tag=target });
                }
                var run = new Button { Text="Run", DialogResult=DialogResult.OK, Enabled=false };
                var cancel = new Button { Text="Cancel", DialogResult=DialogResult.Cancel };
                var buttons = new FlowLayoutPanel { Dock=DockStyle.Bottom, AutoSize=true, FlowDirection=FlowDirection.RightToLeft };
                buttons.Controls.AddRange(new Control[] { cancel, run });
                if (list.Items.Count == 0)
                    picker.Controls.Add(new Label { Dock=DockStyle.Fill, Padding=new Padding(12),
                        Text="No macros found. Add an @macro function to a startup folder and choose Refresh in Startup Folders." });
                else
                    picker.Controls.Add(list);
                picker.Controls.Add(buttons);
                picker.AcceptButton = run;
                picker.CancelButton = cancel;
                list.SelectedIndexChanged += (sender, e) => run.Enabled = list.SelectedItems.Count == 1;
                list.DoubleClick += (sender, e) => run.PerformClick();
                if (picker.ShowDialog(new Owner()) != DialogResult.OK) return;
                var item = list.SelectedItems[0];
                string id = (string)item.Tag;
                var output = new OutputWindow(item.Text);
                output.Show(new Owner());
                output.Run(item.Text, () => {
                    using (var session = new ExcelObjects())
                        return (string)Runtime.Request(new { op="macro", target=id, excelPipe=session.Name });
                });
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
        internal async void Run(string name, Func<string> action)
        {
            output.Text="Running " + name + Environment.NewLine;
            try
            {
                string result = await Task.Run(action);
                if (!IsDisposed) output.AppendText(result.Replace("\r\n","\n").Replace("\n","\r\n") + "\r\nFinished.");
            }
            catch(Exception e) { Runtime.Log(e.Message); if (!IsDisposed) output.AppendText(e.Message); }
        }
    }
}

