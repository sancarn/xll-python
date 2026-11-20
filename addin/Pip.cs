using System;
using System.Diagnostics;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PythonExcel
{
    internal sealed class PipWindow : Form
    {
        readonly TextBox output=Ui.ConsoleBox(), input=new TextBox{Dock=DockStyle.Fill};
        readonly Button run;
        internal PipWindow()
        {
            Text="Python package installer"; ClientSize=new Size(760,460);
            var row=new TableLayoutPanel{Dock=DockStyle.Bottom,Height=34,ColumnCount=3};
            row.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            row.Controls.Add(new Label{Text="pip",AutoSize=true,Margin=new Padding(8,7,4,0)},0,0);
            row.Controls.Add(input,1,0);
            run=Ui.Button("Run",(s,e)=>Execute()); row.Controls.Add(run,2,0);
            Controls.Add(output); Controls.Add(row); AcceptButton=run;
            output.Text=Runtime.Python+"\r\nEnter a pip command, such as list or install openpyxl.\r\n";
            Shown+=(s,e)=>input.Focus();
        }

        internal static string Arguments(string command)
        {
            string args=command.Trim();
            if (args.Equals("pip",StringComparison.OrdinalIgnoreCase)) args="";
            else if (args.StartsWith("pip ",StringComparison.OrdinalIgnoreCase)) args=args.Substring(4);
            return "-I -X utf8 -u -m pip --no-input " + args;
        }

        void Write(string line)
        {
            if (IsDisposed || !IsHandleCreated) return;
            if (InvokeRequired)
            {
                try { BeginInvoke(new Action<string>(Write), line); } catch (InvalidOperationException) {}
            }
            else output.AppendText(line+"\r\n");
        }

        async void Execute()
        {
            string command=input.Text;
            if(command.Trim().Length==0) return;
            input.Clear(); input.Enabled=run.Enabled=false; Write("> pip "+command);
            try
            {
                int code=await Task.Run(()=> {
                    using(var gate=new Mutex(false,@"Local\XllPython.Pip."+System.IO.Path.GetFileName(Runtime.Home)))
                    {
                        try { gate.WaitOne(); } catch(AbandonedMutexException) {}
                        try {
                            using(var p=new Process{StartInfo=Runtime.PipProcess(Arguments(command))})
                            {
                                p.OutputDataReceived+=(s,e)=>{ if(e.Data!=null) Write(e.Data); };
                                p.ErrorDataReceived+=(s,e)=>{ if(e.Data!=null) Write(e.Data); };
                                p.Start(); p.BeginOutputReadLine(); p.BeginErrorReadLine(); p.WaitForExit();
                                return p.ExitCode;
                            }
                        } finally { gate.ReleaseMutex(); }
                    }
                });
                Write("pip exited with code "+code);
            }
            catch(Exception e) { Write(e.Message); }
            finally { if(!IsDisposed) { input.Enabled=run.Enabled=true; input.Focus(); } }
        }
    }
}

