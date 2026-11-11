using System;
using System.Runtime.InteropServices;
using System.Text;

namespace CallFromExcel.AddIn
{
    // C# port of modWindowUtils.bas's window-enumeration helpers (NumWindowsWithCaption,
    // GetHandleFromPartialCaption, WindowTitleFromHandle) -- same Win32 API, same approach
    // (walk top-level windows via GetWindow(GW_HWNDNEXT) starting from FindWindow(null,
    // null), rather than EnumWindows, to match the VBA version's exact behavior).
    public static class WindowUtils
    {
        private const int GW_HWNDNEXT = 2;

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindow(IntPtr hWnd, int uCmd);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        private const int SW_RESTORE = 9;

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        // Restores the window if it's minimised, and brings it to the front. Windows allows the
        // latter because Excel, the caller's process, is in the foreground (the user just
        // clicked its ribbon).
        public static void ShowAndActivate(IntPtr hWnd)
        {
            if (IsIconic(hWnd)) ShowWindow(hWnd, SW_RESTORE);
            SetForegroundWindow(hWnd);
        }

        public static string WindowTitleFromHandle(IntPtr hWnd)
        {
            int length = GetWindowTextLength(hWnd);
            if (length == 0) return "";
            var sb = new StringBuilder(length + 1);
            GetWindowText(hWnd, sb, sb.Capacity);
            return sb.ToString();
        }

        public static int NumWindowsWithCaption(string caption)
        {
            int count = 0;
            IntPtr hWnd = FindWindow(null, null);
            while (hWnd != IntPtr.Zero)
            {
                if (WindowTitleFromHandle(hWnd).IndexOf(caption, StringComparison.Ordinal) >= 0)
                {
                    count++;
                }
                hWnd = GetWindow(hWnd, GW_HWNDNEXT);
            }
            return count;
        }

        public static bool TryGetHandleFromPartialCaption(string caption, out IntPtr hWndFound)
        {
            IntPtr hWnd = FindWindow(null, null);
            while (hWnd != IntPtr.Zero)
            {
                if (WindowTitleFromHandle(hWnd).IndexOf(caption, StringComparison.Ordinal) >= 0)
                {
                    hWndFound = hWnd;
                    return true;
                }
                hWnd = GetWindow(hWnd, GW_HWNDNEXT);
            }
            hWndFound = IntPtr.Zero;
            return false;
        }
    }
}
