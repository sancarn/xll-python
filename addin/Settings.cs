using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace PythonExcel
{
    internal sealed class FolderWindow : Form
    {
        readonly ListBox functions = new ListBox(), packages = new ListBox();
        internal Settings Selection => new Settings { folders=functions.Items.Cast<string>().ToArray(),
                                                      packages=packages.Items.Cast<string>().ToArray() };

        internal FolderWindow()
        {
            Text="Python startup folders"; ClientSize=new Size(660,400); MinimumSize=new Size(500,300);
            StartPosition=FormStartPosition.CenterParent;
            var settings=Runtime.LoadSettings();
            var tabs=new TabControl{Dock=DockStyle.Fill};
            AddPage(tabs,functions,settings.folders,"Worksheet functions","Load top-level .py files from these folders when Excel starts.");
            AddPage(tabs,packages,settings.packages,"Package links","Add compatible site-packages or package parent folders to Python’s search path.");
            var footer=new FlowLayoutPanel{Dock=DockStyle.Bottom,AutoSize=true,FlowDirection=FlowDirection.RightToLeft};
            var save=Ui.Button("Save", (s,e)=> {
                try {
                    Runtime.SaveSettings(Selection);
                    DialogResult=DialogResult.OK;
                } catch (Exception error) { MessageBox.Show(error.Message,Text); }
            });
            footer.Controls.Add(save);
            footer.Controls.Add(Ui.Button("Refresh",(s,e)=>{ DialogResult=DialogResult.Retry; }));
            footer.Controls.Add(Ui.Button("Cancel",(s,e)=>Close()));
            footer.Controls.Add(new Label{Text="Refresh saves these folders and reloads functions now.",AutoSize=true,Margin=new Padding(6,9,6,0)});
            Controls.Add(tabs); Controls.Add(footer); AcceptButton=save;
        }

        static void AddPage(TabControl tabs,ListBox list,string[] paths,string title,string help)
        {
            var page=new TabPage(title){Padding=new Padding(8)};
            list.Dock=DockStyle.Fill; list.HorizontalScrollbar=true; list.Items.AddRange(paths);
            var buttons=new FlowLayoutPanel{Dock=DockStyle.Bottom,AutoSize=true};
            buttons.Controls.Add(Ui.Button("Add folder…",(s,e)=> {
                using (var picker=new FolderBrowserDialog{Description=help})
                    if (picker.ShowDialog()==DialogResult.OK &&
                        !list.Items.Cast<string>().Contains(picker.SelectedPath,StringComparer.OrdinalIgnoreCase))
                        list.Items.Add(picker.SelectedPath);
            }));
            buttons.Controls.Add(Ui.Button("Remove",(s,e)=>{ if(list.SelectedIndex>=0) list.Items.RemoveAt(list.SelectedIndex); }));
            page.Controls.Add(list); page.Controls.Add(buttons);
            page.Controls.Add(new Label{Text=help,Dock=DockStyle.Top,Height=40});
            tabs.TabPages.Add(page);
        }
    }
}

