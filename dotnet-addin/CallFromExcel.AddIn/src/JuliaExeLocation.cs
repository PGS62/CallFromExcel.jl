using System;
using System.IO;

namespace CallFromExcel.AddIn
{
    // C# port of JuliaExeLocation (modMain.bas).
    public static class JuliaExeLocation
    {
        public static string Find(bool useLinux = false)
        {
            // Under WSL, "julia" is resolved by the Linux-side shell against its own PATH,
            // not anything Windows-side - no search needed here, matching the VBA original.
            if (useLinux) return "julia";

            // First search PATH.
            string pathEnv = Environment.GetEnvironmentVariable("PATH") ?? "";
            foreach (string entry in pathEnv.Split(';'))
            {
                // Windows accepts a PATH entry in double quotes, e.g. "C:\Program Files\Foo\bin",
                // but Path.Combine throws on the quote character. Strip them, and skip any entry
                // that still isn't a valid path rather than letting it abort the search.
                string folder = entry.Trim().Trim('"');
                if (folder.Length == 0 || folder.IndexOfAny(Path.GetInvalidPathChars()) >= 0) continue;
                string exeFile = Path.Combine(folder, "julia.exe");
                if (File.Exists(exeFile)) return exeFile;
            }

            // If not found on PATH, search the locations the Windows installer uses by
            // default (if the user accepted defaults), choosing the most recently
            // installed if more than one version is found.
            string chosenExe = null;
            DateTime chosenCreatedDate = DateTime.MinValue;
            string[] parentFolders =
            {
                Path.Combine(Environment.GetEnvironmentVariable("LOCALAPPDATA") ?? "", "Programs"),
                Environment.GetEnvironmentVariable("LOCALAPPDATA") ?? ""
            };

            foreach (string parentFolderName in parentFolders)
            {
                if (!Directory.Exists(parentFolderName)) continue;
                foreach (string childFolder in Directory.GetDirectories(parentFolderName))
                {
                    string folderName = Path.GetFileName(childFolder);
                    if (!folderName.StartsWith("Julia", StringComparison.Ordinal)) continue;
                    string exeFile = Path.Combine(childFolder, "bin", "julia.exe");
                    if (!File.Exists(exeFile)) continue;
                    DateTime createdDate = Directory.GetCreationTime(childFolder);
                    if (createdDate > chosenCreatedDate)
                    {
                        chosenCreatedDate = createdDate;
                        chosenExe = exeFile;
                    }
                }
            }

            if (chosenExe == null)
            {
                throw new InvalidOperationException(
                    "Julia executable not found, after looking on the path and in folders to which Julia " +
                    "is typically installed on Windows. When installing Julia check the \"Add Julia to Path\" option.");
            }
            return chosenExe;
        }
    }
}
