using System;
using System.IO;

namespace CallFromExcel.AddIn
{
    // C# port of JuliaExeLocation (modMain.bas), plus juliaup's location.
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
                // so strip them. Plain concatenation, not Path.Combine, which throws on a quote or
                // other invalid character - File.Exists just returns false for an invalid path.
                string folder = entry.Trim().Trim('"');
                if (folder.Length > 0 && File.Exists(folder + "\\julia.exe")) return folder + "\\julia.exe";
            }

            // juliaup's julia.exe (an app execution alias), which a Julia installed with winget or
            // from the Microsoft Store puts on PATH - but Excel's copy of PATH is fixed when Excel
            // starts, so a Julia installed since then (e.g. after NotFoundError's advice) is found
            // here without restarting Excel.
            string windowsApps = Path.Combine(Environment.GetEnvironmentVariable("LOCALAPPDATA") ?? "", "Microsoft", "WindowsApps", "julia.exe");
            if (File.Exists(windowsApps)) return windowsApps;

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

            if (chosenExe == null) throw NotFoundError();
            return chosenExe;
        }

        // The command julialang.org/downloads gives for installing Julia on Windows (with juliaup,
        // from the Microsoft Store).
        public const string InstallCommand = "winget install --name Julia --id 9NJNWW8PVKMN -e -s msstore";
        public const string DownloadsUrl = "https://julialang.org/downloads/";

        // Public, so test/runtests.jl can check the message on a PC that has Julia.
        public static JuliaCommandException NotFoundError()
        {
            return new JuliaCommandException("Julia isn't installed on this PC, or Excel can't find it: it isn't on the " +
                "path, or in the folders Julia is usually installed to. To install Julia, open Command Prompt and run: " +
                InstallCommand + " - then launch Julia from Excel again: if anything else is needed, Excel will say so. " +
                "Other ways to install Julia are described at " + DownloadsUrl + ".", InstallCommand, DownloadsUrl);
        }
    }
}
