using System;
using System.Diagnostics;
using System.IO;
using System.Threading;

namespace JuliaExcel.AddIn
{
    // C# port of the file-handling helpers in modUtils.bas. Most of modUtils.bas's own
    // helpers (FileExists, ReadTextFile, GetFullTempPath, IsProcessRunning) existed to work
    // around VBA/Scripting.FileSystemObject limitations that don't apply here -- .NET's
    // File/Directory/Path/Process classes already do this directly and more simply, so
    // those aren't separately wrapped; only LocalTemp, SaveTextFile (for its retry
    // behaviour) and the PID-based cleanup logic are ported.
    public static class FileUtils
    {
        private const string SubFolderName = "@JuliaExcel";
        private static string _localTemp;

        // Port of LocalTemp (modUtils.bas): a writable directory for files used to
        // communicate with Julia (port numbers, startup scripts, flag/error files).
        public static string LocalTemp()
        {
            if (_localTemp != null) return _localTemp;
            string dir = Path.Combine(Path.GetTempPath(), SubFolderName);
            Directory.CreateDirectory(dir);
            _localTemp = dir;
            return _localTemp;
        }

        // Port of SaveTextFile (modUtils.bas): retries on transient sharing violations --
        // antivirus/indexing can transiently lock a just-written file for a few
        // milliseconds, the same brief, self-resolving glitch the VBA version retries.
        public static void SaveTextFile(string path, string contents)
        {
            const int maxRetries = 10;
            const int delayMs = 25;
            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                try
                {
                    File.WriteAllText(path, contents);
                    return;
                }
                catch (IOException) when (attempt < maxRetries)
                {
                    Thread.Sleep(delayMs);
                }
            }
        }

        public static string ReadTextFile(string path)
        {
            return File.ReadAllText(path);
        }

        // Port of TryExtractPidFromFilename (modUtils.bas): parses "<Prefix>_<PID>.<ext>".
        public static bool TryExtractPidFromFilename(string fileName, out int pid)
        {
            pid = 0;
            int dotPos = fileName.LastIndexOf('.');
            if (dotPos < 0) return false;
            int underscorePos = fileName.LastIndexOf('_', dotPos - 1 < 0 ? 0 : dotPos - 1);
            if (underscorePos < 0) return false;
            string pidStr = fileName.Substring(underscorePos + 1, dotPos - underscorePos - 1);
            return int.TryParse(pidStr, out pid);
        }

        public static bool IsProcessRunning(int pid)
        {
            try
            {
                using (Process.GetProcessById(pid)) { return true; }
            }
            catch (ArgumentException)
            {
                return false;
            }
        }

        // Port of CleanLocalTemp (modUtils.bas): deletes files belonging to Excel sessions
        // that are no longer running, identified by the PID encoded in their file name.
        public static void CleanLocalTemp()
        {
            foreach (string file in Directory.GetFiles(LocalTemp()))
            {
                string name = Path.GetFileName(file);
                if (TryExtractPidFromFilename(name, out int pid) && !IsProcessRunning(pid))
                {
                    try { File.Delete(file); } catch (IOException) { } catch (UnauthorizedAccessException) { }
                }
            }
        }
    }
}
