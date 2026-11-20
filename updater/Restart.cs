using System;
using System.Diagnostics;
using System.IO;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace PythonExcel.Update
{
    internal static class Restart
    {
        [STAThread]
        static int Main()
        {
            RestartPlan plan = null;
            bool ready = false;
            try
            {
                plan = new JavaScriptSerializer().Deserialize<RestartPlan>(Console.ReadLine());
                string args = plan.Arguments();
                string excel;
                using (var parent = Process.GetProcessById(plan.ParentPid))
                {
                    excel = parent.MainModule.FileName;
                    if (!String.Equals(Path.GetFileName(excel), "EXCEL.EXE", StringComparison.OrdinalIgnoreCase))
                        throw new IOException("The expected Excel instance is unavailable.");
                    // Open the process handle before authorizing Quit, avoiding a PID reuse race.
                    var handle = parent.Handle;
                    Console.WriteLine("READY");
                    Console.Out.Flush();
                    ready = true;
                    if (!parent.WaitForExit(90000))
                        throw new IOException("Excel stayed open. The update was canceled.");
                }

                string backup = plan.Staged + ".previous";
                bool replaced = false;
                try
                {
                    if (RestartPlan.Hash(plan.Staged) != plan.ReleaseHash || RestartPlan.Hash(plan.Target) != plan.OriginalHash)
                        throw new IOException("The downloaded or installed XLL changed during the restart.");
                    File.Replace(plan.Staged, plan.Target, backup);
                    replaced = true;
                    StartExcel(excel, args);
                }
                catch (Exception updateError)
                {
                    if (replaced) File.Replace(backup, plan.Target, null);
                    StartExcel(excel, args);
                    throw new IOException("Excel was reopened with the previous add-in. " + updateError.Message);
                }
                if (File.Exists(backup)) File.Delete(backup);
                return 0;
            }
            catch (Exception e)
            {
                if (ready) MessageBox.Show(e.Message, "Python for Excel update");
                else Console.WriteLine(e.Message);
                return 1;
            }
            finally
            {
                if (ready && plan != null && File.Exists(plan.Staged))
                    try { File.Delete(plan.Staged); } catch (IOException) { }
            }
        }

        static void StartExcel(string exe, string args)
        {
            using (Process.Start(new ProcessStartInfo(exe, args) { UseShellExecute = false })) { }
        }
    }
}
