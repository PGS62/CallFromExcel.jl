using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace CallFromExcel.AddIn
{
    // The Julia logo for the ribbon and our dialogs, taken from the icon in the user's own
    // julia.exe - so the add-in doesn't ship a copy of the logo (whose licence, CC BY-NC-SA,
    // doesn't sit well with MIT). The Microsoft Store's julia.exe on PATH is an App Execution
    // Alias with no icon of its own, so a juliaup installation is looked for too. Null when no
    // julia.exe with an icon is found; callers then fall back to an Office or default icon.
    internal static class JuliaIcon
    {
        private static readonly Lazy<string> ExeWithIcon = new Lazy<string>(FindExeWithIcon);

        // The logo at size x size pixels, or null.
        public static Bitmap Bitmap(int size)
        {
            using (Icon icon = Extract(ExeWithIcon.Value, size))
            {
                return icon?.ToBitmap();
            }
        }

        // The logo as a window icon, or null.
        public static Icon WindowIcon(int size)
        {
            return Extract(ExeWithIcon.Value, size);
        }

        private static string FindExeWithIcon()
        {
            string onPath = null;
            try { onPath = JuliaExeLocation.Find(false); } catch (Exception) { }
            if (HasIcon(onPath)) return onPath;

            string juliaup = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".julia", "juliaup");
            if (!Directory.Exists(juliaup)) return null;
            return Directory.GetDirectories(juliaup, "julia-*")
                .Select(d => Path.Combine(d, "bin", "julia.exe"))
                .Where(File.Exists)
                .OrderByDescending(File.GetLastWriteTimeUtc)
                .FirstOrDefault(HasIcon);
        }

        private static bool HasIcon(string exe)
        {
            using (Icon icon = Extract(exe, 16)) return icon != null;
        }

        // A copy of the exe's first icon at the given size, so the extracted handle can be freed.
        private static Icon Extract(string exe, int size)
        {
            if (string.IsNullOrEmpty(exe) || !File.Exists(exe)) return null;
            var handles = new IntPtr[1];
            if (PrivateExtractIcons(exe, 0, size, size, handles, IntPtr.Zero, 1, 0) == 0 || handles[0] == IntPtr.Zero) return null;
            try
            {
                using (Icon icon = Icon.FromHandle(handles[0])) return (Icon)icon.Clone();
            }
            finally
            {
                DestroyIcon(handles[0]);
            }
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern uint PrivateExtractIcons(string file, int index, int cx, int cy, IntPtr[] icons, IntPtr ids, uint count, uint flags);

        [DllImport("user32.dll")]
        private static extern bool DestroyIcon(IntPtr icon);
    }
}
