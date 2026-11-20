using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;

namespace PythonExcel
{
    internal sealed class Settings
    {
        public string[] folders { get; set; } = new string[0];
        public string[] packages { get; set; } = new string[0];
    }

    internal static class Runtime
    {
        internal static readonly string Root = Path.GetFullPath(Environment.GetEnvironmentVariable("XLL_PYTHON_HOME")
            ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "XllPython"));
        internal static string Home, Support;
        internal static string Python => Path.Combine(Home, "python.exe");
        internal static string SettingsPath => Path.Combine(Root, "settings.json");
        internal static JavaScriptSerializer Json => new JavaScriptSerializer { MaxJsonLength = int.MaxValue, RecursionLimit = 128 };
        private static readonly object logLock = new object();
        private static readonly object initLock = new object();
        private static Call invoke;
        private static Free free;

        [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
        static extern IntPtr LoadLibraryEx(string path, IntPtr reserved, uint flags);
        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        static extern IntPtr GetModuleHandle(string name);
        [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
        static extern IntPtr FindResource(IntPtr module, string name, IntPtr type);
        [DllImport("kernel32", SetLastError = true)]
        static extern IntPtr LoadResource(IntPtr module, IntPtr resource);
        [DllImport("kernel32", SetLastError = true)]
        static extern int SizeofResource(IntPtr module, IntPtr resource);
        [DllImport("kernel32")]
        static extern IntPtr LockResource(IntPtr resource);
        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        static extern uint GetModuleFileName(IntPtr handle, StringBuilder path, int size);
        [DllImport("kernel32", CharSet = CharSet.Ansi, ExactSpelling = true, SetLastError = true)]
        static extern IntPtr GetProcAddress(IntPtr lib, string name);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl, CharSet = CharSet.Unicode)]
        delegate IntPtr Start(string home, string exe, string support);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl, CharSet = CharSet.Unicode)]
        delegate IntPtr Call(string json, out int ok);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        delegate void Free(IntPtr value);

        static T Export<T>(IntPtr lib, string name) where T : class
        {
            var addr = GetProcAddress(lib, name);
            if (addr == IntPtr.Zero) throw new Win32Exception();
            return Marshal.GetDelegateForFunctionPointer(addr, typeof(T)) as T;
        }

        internal static void Log(string message)
        {
            try
            {
                lock (logLock)
                {
                    Directory.CreateDirectory(Root);
                    File.AppendAllText(Path.Combine(Root, "python-excel.log"),
                        DateTime.Now.ToString("s") + "  " + message + Environment.NewLine, Encoding.UTF8);
                }
            }
            catch (IOException) { Trace.WriteLine(message); }
        }

        internal static Settings LoadSettings()
        {
            return File.Exists(SettingsPath) ? Json.Deserialize<Settings>(File.ReadAllText(SettingsPath)) : new Settings();
        }

        internal static void SaveSettings(Settings settings)
        {
            Directory.CreateDirectory(Root);
            string temp = SettingsPath + "." + Guid.NewGuid().ToString("N");
            File.WriteAllText(temp, Json.Serialize(settings), Encoding.UTF8);
            try
            {
                if (File.Exists(SettingsPath)) File.Replace(temp, SettingsPath, null);
                else File.Move(temp, SettingsPath);
            }
            finally { if (File.Exists(temp)) File.Delete(temp); }
        }

        static Stream OpenArchive(string resource)
        {
            if (resource != "runtime")
                return Assembly.GetExecutingAssembly().GetManifestResourceStream(resource + ".zip");
            var module = GetModuleHandle(ExcelDna.Integration.ExcelDnaUtil.XllPath);
            if (module == IntPtr.Zero) throw new Win32Exception();
            var entry = FindResource(module, "PYTHON_RUNTIME", new IntPtr(10));
            if (entry == IntPtr.Zero) throw new Win32Exception();
            var data = LockResource(LoadResource(module, entry));
            if (data == IntPtr.Zero) throw new InvalidDataException("Cannot read the embedded Python runtime.");
            var bytes = new byte[SizeofResource(module, entry)];
            Marshal.Copy(data, bytes, 0, bytes.Length);
            return new MemoryStream(bytes, false);
        }

        static string Unpack(string resource, string hash)
        {
            string dir = Path.Combine(Root, resource + "-" + hash.Substring(0, 16));
            if (Directory.Exists(dir)) return dir;
            string temp = dir + "." + Guid.NewGuid().ToString("N").Substring(0,8);
            Directory.CreateDirectory(temp);
            try
            {
                Log("Installing " + resource + " into " + dir);
                using (var stream = OpenArchive(resource))
                using (var digest = SHA256.Create())
                {
                    string actual = BitConverter.ToString(digest.ComputeHash(stream)).Replace("-", "").ToLowerInvariant();
                    if (actual != hash) throw new InvalidDataException(resource + " payload checksum mismatch");
                    stream.Position = 0;
                    using (var zip = new ZipArchive(stream))
                        foreach (var item in zip.Entries)
                        {
                            string file = Path.GetFullPath(Path.Combine(temp, item.FullName));
                            if (!file.StartsWith(temp + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
                                throw new InvalidDataException(item.FullName);
                            Directory.CreateDirectory(Path.GetDirectoryName(file));
                            using (var source = item.Open())
                            using (var dest = File.Create(file)) source.CopyTo(dest);
                        }
                }
                Directory.Move(temp, dir);
                return dir;
            }
            finally { if (Directory.Exists(temp)) Directory.Delete(temp, true); }
        }

        internal static ProcessStartInfo PipProcess(string arguments)
        {
            var info = new ProcessStartInfo(Python, arguments)
            {
                WorkingDirectory = Home, UseShellExecute = false, CreateNoWindow = true,
                RedirectStandardOutput = true, RedirectStandardError = true,
                StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8
            };
            info.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8";
            info.EnvironmentVariables["PYTHONUTF8"] = "1";
            info.EnvironmentVariables["PYTHONNOUSERSITE"] = "1";
            return info;
        }

        static void Install()
        {
            Directory.CreateDirectory(Root);
            string[] versions;
            using (var reader = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream("versions.txt")))
                versions = reader.ReadToEnd().Split('\n');
            using (var mutex = new Mutex(false, @"Local\XllPython.Install." + versions[0]))
            {
                try { mutex.WaitOne(); } catch (AbandonedMutexException) { }
                try
                {
                    Home = Unpack("runtime", versions[0].Trim());
                    Support = Unpack("support", versions[1].Trim());
                    string ready = Path.Combine(Home, ".pip-ready");
                    if (!File.Exists(ready) || !File.Exists(Path.Combine(Home, "Lib", "site-packages", "pip", "__main__.py")))
                    {
                        Log("Bootstrapping pip from the bundled wheel.");
                        using (var p = Process.Start(PipProcess("-I -X utf8 -m ensurepip --default-pip")))
                        {
                            var stderr = p.StandardError.ReadToEndAsync();
                            string output = p.StandardOutput.ReadToEnd();
                            p.WaitForExit();
                            if (p.ExitCode != 0) throw new IOException("pip bootstrap failed\n" + output + stderr.Result);
                        }
                        File.WriteAllText(ready, "OK");
                    }
                }
                finally { mutex.ReleaseMutex(); }
            }
        }

        static IntPtr Load(string path)
        {
            var lib = LoadLibraryEx(path, IntPtr.Zero, 0x1100); // DLL's own directory and safe default directories
            if (lib == IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error(), "Loading " + path);
            return lib;
        }

        internal static void Ensure()
        {
            lock (initLock)
            {
                if (invoke != null) return;
                Install();
                string dll = Path.Combine(Home, "python313.dll");
                IntPtr existing = GetModuleHandle("python313.dll");
                if (existing != IntPtr.Zero)
                {
                    var path = new StringBuilder(32768);
                    GetModuleFileName(existing, path, path.Capacity);
                    if (!String.Equals(path.ToString(), dll, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("This Excel process already loaded Python from " + path);
                }
                Load(dll);
                var bridge = Load(Path.Combine(Support, "bridge.dll"));
                free = Export<Free>(bridge, "xp_free");
                IntPtr error = Export<Start>(bridge, "xp_start")(Home, Python, Support);
                if (error != IntPtr.Zero) throw new InvalidOperationException(Take(error));
                invoke = Export<Call>(bridge, "xp_call");
            }
        }

        static string Take(IntPtr ptr)
        {
            if (ptr == IntPtr.Zero) throw new OutOfMemoryException();
            try { return Marshal.PtrToStringUni(ptr); }
            finally { free(ptr); }
        }

        internal static object Request(object message)
        {
            Ensure();
            int ok;
            string reply = Take(invoke(Json.Serialize(message), out ok));
            if (ok == 0) throw new InvalidOperationException(reply);
            var decoded = Json.Deserialize<Dictionary<string,object>>(reply);
            if (decoded.ContainsKey("failure")) throw new InvalidOperationException((string)decoded["failure"]);
            return decoded["value"];
        }
    }
}

