using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Win32;

namespace CallFromExcel.AddIn
{
    // A most-recently-used list, newest first, kept per user in the registry so it survives Excel
    // restarts. Used for the ribbon's Julia File box and the Launch Settings project folder.
    internal sealed class RecentList
    {
        private const int MaxItems = 10;
        private const string ValueName = "Recent";
        private readonly string _keyPath;

        public RecentList(string keyPath)
        {
            _keyPath = keyPath;
        }

        public List<string> Items
        {
            get
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(_keyPath))
                {
                    return (key?.GetValue(ValueName) as string[] ?? new string[0]).ToList();
                }
            }
        }

        // Puts an item at the top, removing any earlier copy, and drops the oldest beyond
        // MaxItems. Paths, so compared ignoring case.
        public void Add(string item)
        {
            List<string> items = Items;
            items.RemoveAll(f => string.Equals(f, item, StringComparison.OrdinalIgnoreCase));
            items.Insert(0, item);
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(_keyPath))
            {
                key.SetValue(ValueName, items.Take(MaxItems).ToArray(), RegistryValueKind.MultiString);
            }
        }

        public void Clear()
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(_keyPath))
            {
                key.DeleteValue(ValueName, false);
            }
        }
    }
}
