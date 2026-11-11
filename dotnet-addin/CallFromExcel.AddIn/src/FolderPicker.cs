using System;
using System.Runtime.InteropServices;

namespace CallFromExcel.AddIn
{
    // The modern Windows folder picker (the File Open dialog in pick-folders mode), which .NET
    // Framework's FolderBrowserDialog doesn't offer - that shows the old tree-only dialog.
    internal static class FolderPicker
    {
        private const uint FOS_PICKFOLDERS = 0x20;
        private const uint FOS_FORCEFILESYSTEM = 0x40;
        private const uint SIGDN_FILESYSPATH = 0x80058000;
        private const int ERROR_CANCELLED_HRESULT = unchecked((int)0x800704C7);

        // The chosen folder, or null if the user cancelled. initialFolder is where the dialog
        // opens, when it exists.
        public static string Pick(IntPtr owner, string title, string initialFolder)
        {
            var dialog = (IFileOpenDialog)new FileOpenDialogCoClass();
            try
            {
                dialog.GetOptions(out uint options);
                dialog.SetOptions(options | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM);
                dialog.SetTitle(title);
                if (!string.IsNullOrWhiteSpace(initialFolder) && System.IO.Directory.Exists(initialFolder))
                {
                    Guid shellItemGuid = typeof(IShellItem).GUID;
                    if (SHCreateItemFromParsingName(initialFolder, IntPtr.Zero, ref shellItemGuid, out IShellItem folder) == 0)
                    {
                        dialog.SetFolder(folder);
                    }
                }
                int hr = dialog.Show(owner);
                if (hr == ERROR_CANCELLED_HRESULT) return null;
                Marshal.ThrowExceptionForHR(hr);
                dialog.GetResult(out IShellItem result);
                result.GetDisplayName(SIGDN_FILESYSPATH, out IntPtr pathPtr);
                try { return Marshal.PtrToStringUni(pathPtr); }
                finally { Marshal.FreeCoTaskMem(pathPtr); }
            }
            finally
            {
                Marshal.ReleaseComObject(dialog);
            }
        }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = true)]
        private static extern int SHCreateItemFromParsingName(string path, IntPtr bindContext, ref Guid riid, out IShellItem item);

        [ComImport, Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")]
        private class FileOpenDialogCoClass { }

        // IFileOpenDialog's methods, in vtable order (IModalWindow, then IFileDialog). Only the
        // ones used here have real signatures; the rest are placeholders that keep the order.
        [ComImport, Guid("d57c7288-d4ad-4768-be02-9d969532d960"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IFileOpenDialog
        {
            [PreserveSig] int Show(IntPtr owner);
            void SetFileTypes();
            void SetFileTypeIndex();
            void GetFileTypeIndex();
            void Advise();
            void Unadvise();
            void SetOptions(uint options);
            void GetOptions(out uint options);
            void SetDefaultFolder(IShellItem item);
            void SetFolder(IShellItem item);
            void GetFolder();
            void GetCurrentSelection();
            void SetFileName();
            void GetFileName();
            void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string title);
            void SetOkButtonLabel();
            void SetFileNameLabel();
            void GetResult(out IShellItem item);
        }

        [ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IShellItem
        {
            void BindToHandler();
            void GetParent();
            void GetDisplayName(uint sigdnName, out IntPtr name);
            void GetAttributes();
            void Compare();
        }
    }
}
