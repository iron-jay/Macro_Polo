using System;
using System.Collections.Generic;
using Microsoft.Win32;

namespace Macro_Polo.Core
{
    /// <summary>Reads the live machine registry.</summary>
    public sealed class WindowsRegistryValueSource : IRegistryValueSource
    {
        public object GetValue(RegistryRoot root, string subKeyPath, string valueName)
        {
            try
            {
                using (RegistryKey baseKey = OpenBaseKey(root))
                using (RegistryKey key = baseKey.OpenSubKey(subKeyPath))
                {
                    return key == null ? null : key.GetValue(valueName);
                }
            }
            catch (Exception ex)
            {
                Log.Warn("Failed reading " + root + "\\" + subKeyPath + "\\" + valueName, ex);
                return null;
            }
        }

        public IEnumerable<string> GetSubKeyNames(RegistryRoot root, string subKeyPath)
        {
            try
            {
                using (RegistryKey baseKey = OpenBaseKey(root))
                using (RegistryKey key = baseKey.OpenSubKey(subKeyPath))
                {
                    return key == null ? new string[0] : key.GetSubKeyNames();
                }
            }
            catch (Exception ex)
            {
                Log.Warn("Failed enumerating " + root + "\\" + subKeyPath, ex);
                return new string[0];
            }
        }

        public IEnumerable<string> GetValueNames(RegistryRoot root, string subKeyPath)
        {
            try
            {
                using (RegistryKey baseKey = OpenBaseKey(root))
                using (RegistryKey key = baseKey.OpenSubKey(subKeyPath))
                {
                    return key == null ? new string[0] : key.GetValueNames();
                }
            }
            catch (Exception ex)
            {
                Log.Warn("Failed listing values under " + root + "\\" + subKeyPath, ex);
                return new string[0];
            }
        }

        /// <summary>
        /// Opens the 64-bit view on a 64-bit OS. The add-in is loaded in-process by Office, so it
        /// inherits Office's bitness; a 32-bit Office would otherwise be silently redirected into
        /// Wow6432Node and miss machine-wide policy.
        /// </summary>
        private static RegistryKey OpenBaseKey(RegistryRoot root)
        {
            RegistryView view = Environment.Is64BitOperatingSystem
                ? RegistryView.Registry64
                : RegistryView.Default;

            return RegistryKey.OpenBaseKey(
                root == RegistryRoot.LocalMachine ? RegistryHive.LocalMachine : RegistryHive.CurrentUser,
                view);
        }
    }
}
