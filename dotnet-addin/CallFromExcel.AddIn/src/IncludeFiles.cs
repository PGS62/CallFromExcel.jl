using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Win32;

namespace CallFromExcel.AddIn
{
    // The ribbon's Julia File box: the file to include, and the files most recently chosen or included,
    // newest first. Kept per user in the registry, beside the launch settings, so they survive
    // Excel restarts.
    internal static class IncludeFiles
    {
        private const string KeyPath = @"Software\CallFromExcel\IncludeFiles";
        private static readonly RecentList RecentFiles = new RecentList(KeyPath);

        public static string Current
        {
            get
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(KeyPath))
                {
                    return key?.GetValue("Current") as string ?? "";
                }
            }
            set
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(KeyPath))
                {
                    key.SetValue("Current", Clean(value));
                }
            }
        }

        public static List<string> Recent => RecentFiles.Items;

        // Puts a file at the top of the recent list: one just chosen (Choose) or included.
        public static void AddRecent(string fileName) => RecentFiles.Add(fileName);

        // Makes path the current file, typed, pasted, picked from the list or browsed to; and puts
        // it on the recent list if the file exists, so a mistyped path doesn't get there. (Office
        // reports a typed path only on Enter or leaving the box, not keystroke by keystroke.)
        public static void Choose(string path)
        {
            Current = path;
            string chosen = Current;   // as cleaned by the setter
            if (chosen != "" && System.IO.File.Exists(chosen)) AddRecent(chosen);
        }

        public static void ClearRecent() => RecentFiles.Clear();

        // A typed or pasted path, trimmed, without the quotes Explorer's "Copy as path" adds.
        public static string Clean(string path)
        {
            return (path ?? "").Trim().Trim('"').Trim();
        }
    }
}
