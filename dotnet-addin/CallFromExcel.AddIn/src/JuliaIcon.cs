using System;
using System.Drawing;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

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

        // The window icons, kept for the life of the add-in: a window keeps using the handle it's
        // given, so it mustn't be freed.
        private static readonly Lazy<Icon> SmallWindowIcon = new Lazy<Icon>(() => WindowIcon(System.Windows.Forms.SystemInformation.SmallIconSize.Width));
        private static readonly Lazy<Icon> LargeWindowIcon = new Lazy<Icon>(() => WindowIcon(System.Windows.Forms.SystemInformation.IconSize.Width));

        // Gives the console window of a process just started by conhost.exe (conhostPid) the
        // Julia icon, in its title bar and taskbar button. True once done, or if there's no icon
        // to set; false if the window isn't there yet, to try again shortly. A console window
        // reports as its process the program conhost runs in it (julia.exe, or wsl.exe under
        // WSL), not conhost itself - so it's found among the windows of conhost's children.
        // Done here rather than by Julia (CallFromExcel.set_console_icon) so the icon appears as
        // soon as the window does, not seconds later once Julia has loaded its packages.
        // The icon is cosmetic, so any failure finding it or the window (e.g. a folder that can't
        // be read, which Lazy would then rethrow on every call) gives up on it - true - rather
        // than failing the launch it's called from.
        public static bool TrySetOnConsole(int conhostPid)
        {
            try
            {
                return SetOnConsole(conhostPid);
            }
            catch (Exception)
            {
                return true;
            }
        }

        private static bool SetOnConsole(int conhostPid)
        {
            Icon small = SmallWindowIcon.Value, large = LargeWindowIcon.Value;
            if (small == null || large == null) return true;
            HashSet<uint> children = ChildProcessIds((uint)conhostPid);
            if (children.Count == 0) return false;
            IntPtr found = IntPtr.Zero;
            var className = new StringBuilder(64);
            EnumWindows((hwnd, lParam) =>
            {
                className.Clear();
                GetClassName(hwnd, className, className.Capacity);
                if (className.ToString() != "ConsoleWindowClass") return true;
                GetWindowThreadProcessId(hwnd, out uint pid);
                if (!children.Contains(pid)) return true;
                found = hwnd;
                return false;
            }, IntPtr.Zero);
            if (found == IntPtr.Zero) return false;
            const uint WM_SETICON = 0x0080;
            SendMessage(found, WM_SETICON, (IntPtr)0, small.Handle);   // ICON_SMALL
            SendMessage(found, WM_SETICON, (IntPtr)1, large.Handle);   // ICON_BIG
            return true;
        }

        private static HashSet<uint> ChildProcessIds(uint parentPid)
        {
            var children = new HashSet<uint>();
            IntPtr snapshot = CreateToolhelp32Snapshot(0x2, 0);   // TH32CS_SNAPPROCESS
            if (snapshot == (IntPtr)(-1)) return children;
            try
            {
                var entry = new PROCESSENTRY32 { dwSize = (uint)Marshal.SizeOf(typeof(PROCESSENTRY32)) };
                for (bool more = Process32First(snapshot, ref entry); more; more = Process32Next(snapshot, ref entry))
                {
                    if (entry.th32ParentProcessID == parentPid) children.Add(entry.th32ProcessID);
                }
            }
            finally
            {
                CloseHandle(snapshot);
            }
            return children;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct PROCESSENTRY32
        {
            public uint dwSize;
            public uint cntUsage;
            public uint th32ProcessID;
            public IntPtr th32DefaultHeapID;
            public uint th32ModuleID;
            public uint cntThreads;
            public uint th32ParentProcessID;
            public int pcPriClassBase;
            public uint dwFlags;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string szExeFile;
        }

        private delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr CreateToolhelp32Snapshot(uint flags, uint processId);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "Process32FirstW")]
        private static extern bool Process32First(IntPtr snapshot, ref PROCESSENTRY32 entry);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, EntryPoint = "Process32NextW")]
        private static extern bool Process32Next(IntPtr snapshot, ref PROCESSENTRY32 entry);

        [DllImport("kernel32.dll")]
        private static extern bool CloseHandle(IntPtr handle);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hwnd, StringBuilder className, int maxCount);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam);

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
