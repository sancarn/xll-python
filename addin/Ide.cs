using System;
using System.Collections;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace PythonExcel
{
    internal sealed class IdeWindow : Form
    {
        static readonly object lifetime=new object();
        static IdeWindow current;
        static Thread thread;
        static bool loaderConfigured;
        readonly WebView2 web=new WebView2{Dock=DockStyle.Fill};
        readonly TextBox output=Ui.ConsoleBox(), formula=new TextBox{Width=300};
        readonly Label file=new Label{AutoSize=true,Margin=new Padding(8,8,4,0)};
        readonly NumericUpDown line=new NumericUpDown{Width=65,Minimum=0,Maximum=10000000};
        readonly System.Windows.Forms.Timer timer=new System.Windows.Forms.Timer{Interval=250};
        string target, lastState;
        bool running;

        internal static void Open()
        {
            Runtime.Ensure();
            lock(lifetime)
            {
                if(current!=null && current.IsHandleCreated) { current.BeginInvoke(new Action(current.Activate)); return; }
                if(thread!=null) return;
                thread=new Thread(()=> {
                    try
                    {
                        using(var window=new IdeWindow())
                        {
                            lock(lifetime) current=window;
                            Application.Run(window);
                        }
                    }
                    catch(Exception e) { Runtime.Log(e.ToString()); }
                    finally { lock(lifetime) { current=null; thread=null; } }
                }) { IsBackground=true,Name="Excel Python IDE" };
                thread.SetApartmentState(ApartmentState.STA); thread.Start();
            }
        }

        internal static void Shutdown()
        {
            lock(lifetime)
                if(current!=null && current.IsHandleCreated) current.BeginInvoke(new Action(current.Close));
        }

        IdeWindow()
        {
            Text="Python IDE — Excel"; ClientSize=new Size(1140,780); MinimumSize=new Size(760,500);
            var scriptBar=new FlowLayoutPanel{Dock=DockStyle.Top,AutoSize=true};
            scriptBar.Controls.Add(Ui.Button("Choose script…",(s,e)=>Choose()));
            scriptBar.Controls.Add(Ui.Button("Run saved script",(s,e)=>RunScript()));
            scriptBar.Controls.Add(file);
            file.Text="Open a local Python file in VS Code or choose a script.";
            var debugBar=new FlowLayoutPanel{Dock=DockStyle.Top,AutoSize=true};
            debugBar.Controls.Add(new Label{Text="Breakpoint line (0 = first)",AutoSize=true,Margin=new Padding(6,8,0,0)});
            debugBar.Controls.Add(line);
            debugBar.Controls.Add(Ui.Button("Arm breakpoint",(s,e)=>Arm()));
            foreach(string cmd in new[]{"Step","Continue","Detach"})
            {
                string action=cmd=="Detach"?"stop":cmd.ToLowerInvariant();
                debugBar.Controls.Add(Ui.Button(cmd,(s,e)=>DebugCommand(action)));
            }
            var formulaBar=new FlowLayoutPanel{Dock=DockStyle.Bottom,AutoSize=true};
            formulaBar.Controls.Add(new Label{Text="Excel formula",AutoSize=true,Margin=new Padding(6,8,0,0)});
            formulaBar.Controls.Add(formula);
            formulaBar.Controls.Add(Ui.Button("Evaluate",(s,e)=>Evaluate()));
            formula.KeyDown+=(s,e)=>{if(e.KeyCode==Keys.Enter){e.SuppressKeyPress=true;Evaluate();}};
            var split=new SplitContainer{Dock=DockStyle.Fill,Orientation=Orientation.Horizontal,Height=640,SplitterDistance=495};
            split.Panel1.Controls.Add(web); split.Panel2.Controls.Add(output);
            Controls.Add(split); Controls.Add(formulaBar); Controls.Add(debugBar); Controls.Add(scriptBar);
            Shown+=async(s,e)=>await StartEditor();
            timer.Tick+=(s,e)=>Poll();
            FormClosed+=(s,e)=>{ timer.Stop(); DebugCommand("stop"); timer.Dispose(); };
        }

        // Transfer only handles selected by the user in the native picker. Web content
        // cannot supply filesystem paths, invoke the interpreter, or evaluate formulas.
        const string TrackSelection=@"
(() => {
 if (window.top !== window || location.origin !== 'https://vscode.dev') return;
 let pending;
 window.chrome.webview.addEventListener('message', e => {
   if (!pending || e.data?.type !== 'file-result') return;
   const {resolve,reject,save}=pending; pending=undefined;
   const handles=Array.from(e.additionalObjects || []);
   if (!handles.length) reject(new DOMException('File selection cancelled','AbortError'));
   else resolve(save ? handles[0] : handles);
 });
 for (const [name,save] of [['showOpenFilePicker',false],['showSaveFilePicker',true]]) {
   window[name]=async (options={}) => {
     if (!navigator.userActivation.isActive || pending)
       throw new DOMException('Choose a file from the editor','SecurityError');
     return new Promise((resolve,reject) => {
       pending={resolve,reject,save};
       window.chrome.webview.postMessage({type:'file-picker',save,multiple:!!options.multiple,
                                          suggestion:String(options.suggestedName || '')});
     });
   };
 }
})();";

        static bool EditorOrigin(string url)
        {
            return Uri.TryCreate(url,UriKind.Absolute,out var uri)
                && uri.GetLeftPart(UriPartial.Authority)=="https://vscode.dev";
        }

        void FileRequest(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            var core=web.CoreWebView2;
            if(!EditorOrigin(e.Source) || !EditorOrigin(core.Source)) return;
            if(e.WebMessageAsJson.Length>4096) return;
            var request=Runtime.Json.DeserializeObject(e.WebMessageAsJson) as IDictionary;
            if(request==null || !Equals(request["type"],"file-picker")) return;
            BeginInvoke(new Action(()=>PickForEditor(core,request)));
        }

        void PickForEditor(CoreWebView2 core, IDictionary request)
        {
            if(!EditorOrigin(core.Source)) return;
            var handles=new System.Collections.Generic.List<object>();
            string[] selected=new string[0];
            try
            {
                if(Equals(request["save"],true))
                {
                    using(var dialog=new SaveFileDialog{Title="Save file for editing in vscode.dev",
                        Filter="All files (*.*)|*.*",RestoreDirectory=true,
                        FileName=Path.GetFileName(Convert.ToString(request["suggestion"]))})
                    {
                        if(dialog.ShowDialog(this)==DialogResult.OK)
                        {
                            if(!File.Exists(dialog.FileName)) using(File.Open(dialog.FileName,FileMode.CreateNew)) {}
                            selected=new[]{dialog.FileName};
                        }
                    }
                }
                else
                {
                    using(var dialog=new OpenFileDialog{Title="Open file for editing in vscode.dev",
                        Filter="All files (*.*)|*.*",RestoreDirectory=true,Multiselect=Equals(request["multiple"],true)})
                        if(dialog.ShowDialog(this)==DialogResult.OK) selected=dialog.FileNames;
                }
                if(!EditorOrigin(core.Source)) return;
                foreach(string path in selected)
                    handles.Add(core.Environment.CreateWebFileSystemFileHandle(path,CoreWebView2FileSystemHandlePermission.ReadWrite));
                if(selected.Length>0 && (selected[0].EndsWith(".py",StringComparison.OrdinalIgnoreCase) ||
                   selected[0].EndsWith(".pyw",StringComparison.OrdinalIgnoreCase))) SelectScript(selected[0]);
            }
            catch(Exception error){handles.Clear(); Print(error.Message);}
            if(EditorOrigin(core.Source)) core.PostWebMessageAsJson(@"{""type"":""file-result""}",handles);
        }

        async Task StartEditor()
        {
            try
            {
                if(!loaderConfigured)
                {
                    CoreWebView2Environment.SetLoaderDllFolderPath(Runtime.Support);
                    loaderConfigured=true;
                }
                web.CreationProperties=new CoreWebView2CreationProperties{UserDataFolder=Path.Combine(Runtime.Root,"WebView2")};
                await web.EnsureCoreWebView2Async();
                web.CoreWebView2.WebMessageReceived+=FileRequest;
                await web.CoreWebView2.AddScriptToExecuteOnDocumentCreatedAsync(TrackSelection);
                web.CoreWebView2.NavigationCompleted+=(s,e)=> {
                    if(!e.IsSuccess) Print("VS Code navigation: "+e.WebErrorStatus);
                };
                web.Source=new Uri("https://vscode.dev/");
            }
            catch(Exception e){ Print(e.Message); Runtime.Log(e.ToString()); }
        }

        void SelectScript(string path){target=path;file.Text=path;}
        bool Choose()
        {
            using(var picker=Ui.ScriptPicker())
            {
                if(picker.ShowDialog(this)!=DialogResult.OK) return false;
                SelectScript(picker.FileName); return true;
            }
        }
        void Print(string text)
        {
            if(!IsDisposed) output.AppendText(text.Replace("\r\n","\n").Replace("\n","\r\n")+"\r\n");
        }
        async void RunScript()
        {
            if(running || (target==null && !Choose())) return;
            string path=target;
            running=true; Print("Running "+path);
            try { Print((string)await Task.Run(()=>Runtime.Request(new {op="script",path}))); Print("Finished."); }
            catch(Exception e){Print(e.Message);}
            finally{running=false;}
        }

        void Arm()
        {
            if(target==null && !Choose()) return;
            try {
                Runtime.Request(new {op="arm",path=target,line=(int)line.Value});
                lastState=null; timer.Start(); Print("Breakpoint armed for "+target+". Run the script or recalculate its worksheet function.");
            } catch(Exception e){Print(e.Message);}
        }

        void DebugCommand(string command)
        {
            try { Runtime.Request(new {op="debug",command}); } catch(Exception e){Print(e.Message);}
        }
        void Poll()
        {
            try {
                var state=(IDictionary)Runtime.Request(new {op="debug"});
                string status=(string)state["status"], signature=Runtime.Json.Serialize(state);
                if(signature==lastState) return;
                lastState=signature;
                if(status=="paused") Print(state["function"]+" line "+state["line"]+"    "+state["source"]+"\r\n"+Runtime.Json.Serialize(state["locals"]));
                else Print("Debugger "+status);
                if(status=="finished" || status=="detached") timer.Stop();
            } catch(Exception e){timer.Stop();Print(e.Message);}
        }

        async void Evaluate()
        {
            string expression=formula.Text.Trim();
            if(expression.Length==0) return;
            try
            {
                string result=null;
                await ExcelAsyncUtil.QueueAsMacroTask(()=> {
                    dynamic app=ExcelDnaUtil.Application;
                    object value=app.Evaluate(expression);
                    if (value!=null && System.Runtime.InteropServices.Marshal.IsComObject(value))
                        value=((dynamic)value).Value2;
                    var values=value as Array;
                    if(values!=null && values.Rank==2)
                    {
                        var rows=new object[values.GetLength(0)][];
                        for(int r=0;r<rows.Length;r++)
                            rows[r]=Enumerable.Range(0,values.GetLength(1)).Select(c=>values.GetValue(r+values.GetLowerBound(0),c+values.GetLowerBound(1))).ToArray();
                        result=Runtime.Json.Serialize(rows);
                    }
                    else result=Convert.ToString(value,System.Globalization.CultureInfo.InvariantCulture);
                });
                Print(expression+" = "+result);
            }
            catch(Exception e){Print(e.Message);}
        }
    }
}

