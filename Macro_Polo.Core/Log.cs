using System;
using System.Globalization;
using System.IO;
using Microsoft.Win32;

namespace Macro_Polo.Core
{
    /// <summary>
    /// Minimal opt-in file log. An add-in that throws during startup gets disabled by Office and
    /// the user is given no useful explanation, so the add-in swallows its own errors — this is
    /// how support gets to see them afterwards.
    /// </summary>
    /// <remarks>
    /// Enable by setting <c>HKCU\Software\Macro_Polo\Logging</c> (DWORD) to 1. Output goes to
    /// <c>%LOCALAPPDATA%\Macro Polo\macro-polo.log</c>. Nothing here is allowed to throw.
    /// </remarks>
    public static class Log
    {
        private static readonly object Gate = new object();
        private static string _path;

        /// <summary>
        /// Deliberately re-read rather than cached. Caching it meant the answer was fixed by
        /// whatever the registry said at the moment of the very first log call in the process -
        /// so turning logging on while Office was already running produced a log that was missing
        /// its own opening entries, which is exactly the situation where the log is needed.
        /// </summary>
        public static bool IsEnabled
        {
            get { return ReadEnabledFlag(); }
        }

        public static void Info(string message)
        {
            Write("INFO ", message, null);
        }

        public static void Warn(string message, Exception exception)
        {
            Write("WARN ", message, exception);
        }

        public static void Error(string message, Exception exception)
        {
            Write("ERROR", message, exception);
        }

        private static void Write(string level, string message, Exception exception)
        {
            if (!IsEnabled)
            {
                return;
            }

            try
            {
                string line = string.Format(
                    CultureInfo.InvariantCulture,
                    "{0:yyyy-MM-dd HH:mm:ss.fff} [{1}] {2}{3}",
                    DateTime.Now,
                    level,
                    message,
                    exception == null ? string.Empty : Environment.NewLine + exception);

                lock (Gate)
                {
                    string path = GetPath();
                    if (path != null)
                    {
                        File.AppendAllText(path, line + Environment.NewLine);
                    }
                }
            }
            catch
            {
                // A logger that can break the add-in is worse than no logger.
            }
        }

        private static string GetPath()
        {
            if (_path == null)
            {
                string folder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Macro Polo");

                Directory.CreateDirectory(folder);
                _path = Path.Combine(folder, "macro-polo.log");
            }

            return _path;
        }

        private static bool ReadEnabledFlag()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Macro_Polo"))
                {
                    object value = key == null ? null : key.GetValue("Logging");
                    return value != null && Convert.ToInt32(value, CultureInfo.InvariantCulture) != 0;
                }
            }
            catch
            {
                return false;
            }
        }
    }
}
