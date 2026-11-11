using System;
using Microsoft.Win32;

namespace CallFromExcel.AddIn
{
    // The ribbon's Launch Julia settings, kept per user in the registry so they survive Excel
    // restarts and reinstalls. Only the ribbon uses them: a JuliaLaunch() formula always uses its
    // own arguments, so a workbook launches Julia the same way on every machine.
    public sealed class LaunchSettings
    {
        private const string KeyPath = @"Software\CallFromExcel\LaunchSettings";

        public bool UseLinux;
        public bool MinimiseWindow;
        public string ProjectFolder = "";
        public string OtherOptions = "";
        public string Packages = "";
        public string BashStatements = "";
        public int TimeOut = 30;

        // The project folder and other options combined, as JuliaLaunch's CommandLineOptions.
        public string CommandLineOptions
        {
            get { return (JuliaLaunch.ProjectOption(ProjectFolder, UseLinux) + " " + OtherOptions.Trim()).Trim(); }
        }

        public static LaunchSettings Load()
        {
            var s = new LaunchSettings();
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(KeyPath))
            {
                if (key == null) return s;
                s.UseLinux = ReadInt(key, "UseLinux", 0) != 0;
                s.MinimiseWindow = ReadInt(key, "MinimiseWindow", 0) != 0;
                s.ProjectFolder = ReadString(key, "ProjectFolder");
                s.OtherOptions = ReadString(key, "OtherOptions");
                s.Packages = ReadString(key, "Packages");
                s.BashStatements = ReadString(key, "BashStatements");
                s.TimeOut = Math.Max(1, ReadInt(key, "TimeOut", 30));
            }
            return s;
        }

        public void Save()
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(KeyPath))
            {
                key.SetValue("UseLinux", UseLinux ? 1 : 0, RegistryValueKind.DWord);
                key.SetValue("MinimiseWindow", MinimiseWindow ? 1 : 0, RegistryValueKind.DWord);
                key.SetValue("ProjectFolder", ProjectFolder ?? "");
                key.SetValue("OtherOptions", OtherOptions ?? "");
                key.SetValue("Packages", Packages ?? "");
                key.SetValue("BashStatements", BashStatements ?? "");
                key.SetValue("TimeOut", TimeOut, RegistryValueKind.DWord);
            }
        }

        private static string ReadString(RegistryKey key, string name)
        {
            return key.GetValue(name) as string ?? "";
        }

        private static int ReadInt(RegistryKey key, string name, int fallback)
        {
            object value = key.GetValue(name);
            return value is int i ? i : fallback;
        }
    }
}
