using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace PythonExcel
{
    // The queue and its transcript live with Excel, independently of the WebView.
    internal static class Pip
    {
        sealed class Job
        {
            internal int id, revision;
            internal string command, package, state = "queued";
            internal int? exitCode;
            internal readonly StringBuilder output = new StringBuilder();
        }

        sealed class Distribution
        {
            public string name { get; set; }
            public string version { get; set; }
        }

        static readonly object sync = new object();
        static readonly List<Job> jobs = new List<Job>();
        static Task tail = Task.CompletedTask;
        static Distribution[] installed;
        static string inventoryError;
        static int revision, inventoryRevision;

        static string Key(string name) => Regex.Replace(name.ToLowerInvariant(), @"[-_.]+", "-");

        internal static void Submit(string command, string package)
        {
            if (package != null)
            {
                if (!Regex.IsMatch(package, @"\A[A-Za-z0-9](?:[A-Za-z0-9._-]*[A-Za-z0-9])?\z"))
                    throw new ArgumentException("Invalid package name.");
                command = "install " + package;
            }
            else
            {
                command = command.Trim();
                if (command.Length == 0) return;
                command = Regex.Replace(command, @"^pip(?:\s+|$)", "", RegexOptions.IgnoreCase);
            }
            lock (sync)
            {
                if (package != null && jobs.Any(j => j.package != null && Key(j.package) == Key(package)
                    && (j.state == "queued" || j.state == "running"))) return;
                var job = new Job { id = jobs.Count + 1, revision = ++revision, command = command, package = package };
                jobs.Add(job);
                Schedule(job);
            }
        }

        internal static void Refresh() { lock (sync) Schedule(null); }

        static void Schedule(Job job)
        {
            tail = tail.ContinueWith(previous => Execute(job), TaskScheduler.Default);
        }

        static ProcessStartInfo StartInfo(string command) =>
            Runtime.PipProcess("-I -X utf8 -u -m pip --no-input --disable-pip-version-check " + command);

        static void Append(Job job, string line)
        {
            if (line == null) return;
            lock (sync)
            {
                job.output.AppendLine(line);
                const int limit = 128 * 1024;
                if (job.output.Length > limit)
                {
                    job.output.Remove(0, job.output.Length - limit * 3 / 4);
                    job.output.Insert(0, "[Earlier output omitted]\n");
                }
                job.revision = ++revision;
            }
        }

        static int Run(Job job)
        {
            lock (sync) { job.state = "running"; job.revision = ++revision; }
            using (var process = new Process { StartInfo = StartInfo(job.command) })
            {
                process.OutputDataReceived += (s,e) => Append(job, e.Data);
                process.ErrorDataReceived += (s,e) => Append(job, e.Data);
                process.Start();
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();
                process.WaitForExit();
                return process.ExitCode;
            }
        }

        static Distribution[] ReadInstalled()
        {
            using (var process = Process.Start(StartInfo("list --format=json")))
            {
                var error = process.StandardError.ReadToEndAsync();
                string json = process.StandardOutput.ReadToEnd();
                process.WaitForExit();
                if (process.ExitCode != 0) throw new InvalidOperationException(error.Result.Trim());
                return Runtime.Json.Deserialize<Distribution[]>(json);
            }
        }

        static void Execute(Job job)
        {
            int? code = null;
            Distribution[] inventory = null;
            string error = null;
            try
            {
                using (var gate = new Mutex(false, @"Local\XllPython.Pip." + System.IO.Path.GetFileName(Runtime.Home)))
                {
                    try { gate.WaitOne(); } catch (AbandonedMutexException) {}
                    try
                    {
                        if (job != null) code = Run(job);
                        inventory = ReadInstalled();
                    }
                    finally { gate.ReleaseMutex(); }
                }
            }
            catch (Exception ex) { error = "Could not read the package environment. " + ex.Message; }
            lock (sync)
            {
                installed = inventory;
                inventoryError = error;
                inventoryRevision = ++revision;
                if (job == null) return;
                bool verified = job.package == null || (inventory != null && inventory.Any(p => Key(p.name) == Key(job.package)));
                if (error != null) Append(job, error);
                else if (code == 0 && !verified) Append(job, "Package was not found in this Python environment after installation.");
                job.exitCode = code;
                job.state = code == 0 && verified ? "done" : "failed";
                job.revision = ++revision;
            }
        }

        internal static string Changes(ref int since)
        {
            lock (sync)
            {
                if (since == revision) return null;
                int cursor = since;
                var changed = jobs.Where(j => j.revision > cursor).Select(j => new {
                    j.id, j.command, j.package, j.state, j.exitCode, output = j.output.ToString()
                }).ToArray();
                var inventory = inventoryRevision > cursor ? new { installed, error = inventoryError } : null;
                since = revision;
                return Runtime.Json.Serialize(new { jobs = changed, inventory, python = Runtime.Python });
            }
        }
    }
}
