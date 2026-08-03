using System;

namespace Macro_Polo.Core
{
    /// <summary>
    /// Turns the version string reported by the Office object model into the "{major}.0" form
    /// that Office's own registry paths use.
    /// </summary>
    public static class OfficeVersion
    {
        /// <summary>Used when the host reports nothing usable. 16.0 covers Office 2016 onwards.</summary>
        public const string Fallback = "16.0";

        /// <summary>
        /// Normalises a version such as "16.0" or "16.0.17928.20114" to "16.0". Office keys its
        /// settings by major version only, and hard-coding a single version silently breaks the
        /// lookup on every other release.
        /// </summary>
        public static string Normalize(string reportedVersion)
        {
            if (string.IsNullOrEmpty(reportedVersion))
            {
                return Fallback;
            }

            string[] parts = reportedVersion.Split('.');
            int major;

            if (!int.TryParse(parts[0], out major) || major <= 0)
            {
                return Fallback;
            }

            return major.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".0";
        }

        /// <summary>
        /// Reads <c>Application.Version</c> off the host without taking a compile-time dependency
        /// on either interop assembly, and without letting a COM failure take the add-in down.
        /// </summary>
        public static string FromHost(object application)
        {
            if (application == null)
            {
                return Fallback;
            }

            try
            {
                object version = application.GetType().InvokeMember(
                    "Version",
                    System.Reflection.BindingFlags.GetProperty,
                    null,
                    application,
                    null,
                    System.Globalization.CultureInfo.InvariantCulture);

                return Normalize(version as string);
            }
            catch (Exception ex)
            {
                Log.Warn("Could not read Application.Version; assuming " + Fallback, ex);
                return Fallback;
            }
        }
    }
}
