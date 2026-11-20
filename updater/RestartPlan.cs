using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;

namespace PythonExcel.Update
{
    // Sent over the child process's stdin so workbook paths never enter a shell.
    public sealed class RestartPlan
    {
        public int ParentPid { get; set; }
        public string Target { get; set; }
        public string Staged { get; set; }
        public string OriginalHash { get; set; }
        public string ReleaseHash { get; set; }
        public string[] Files { get; set; }

        public static string Hash(string path)
        {
            using (var hash = SHA256.Create())
            using (var stream = File.OpenRead(path))
                return BitConverter.ToString(hash.ComputeHash(stream)).Replace("-", "").ToLowerInvariant();
        }

        public string Arguments()
        {
            if (Files.Any(s => String.IsNullOrEmpty(s) || s.Contains("\"") || s.EndsWith("\\")))
                throw new IOException("A workbook path cannot be passed to Excel.");
            string args = "/x " + String.Join(" ", Files.Select(s => "\"" + s + "\""));
            if (args.Length > 32000) throw new IOException("The workbook paths exceed the Windows command-line limit.");
            return args;
        }
    }

}
