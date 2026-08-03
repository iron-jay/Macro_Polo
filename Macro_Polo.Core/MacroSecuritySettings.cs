using System;
using System.Collections.Generic;

namespace Macro_Polo.Core
{
    /// <summary>
    /// The effective Trust Center configuration for one Office application, after the various
    /// registry hives have been resolved against each other.
    /// </summary>
    public sealed class MacroSecuritySettings
    {
        public MacroSecuritySettings()
        {
            TrustedLocations = new List<TrustedLocation>();
        }

        /// <summary>The effective macro setting.</summary>
        public VbaWarningLevel WarningLevel { get; set; }

        /// <summary>
        /// True when <see cref="WarningLevel"/> came from a Group Policy hive rather than from the
        /// user's own preferences, meaning the user cannot change it in the Trust Center.
        /// </summary>
        public bool IsManagedByPolicy { get; set; }

        /// <summary>
        /// The "Block macros from running in Office files from the Internet" policy
        /// (<c>blockcontentexecutionfrominternet</c>).
        /// </summary>
        public bool BlockMacrosFromInternet { get; set; }

        /// <summary>Trusted Locations configured for this application.</summary>
        public List<TrustedLocation> TrustedLocations { get; private set; }

        /// <summary>
        /// The "Disable all Trusted Locations" setting. When true, the entries in
        /// <see cref="TrustedLocations"/> are ignored by Office.
        /// </summary>
        public bool AllTrustedLocationsDisabled { get; set; }

        /// <summary>
        /// The macro setting as Office will actually apply it. Office treats a missing value as
        /// <see cref="VbaWarningLevel.DisableWithNotification"/>.
        /// </summary>
        public VbaWarningLevel EffectiveWarningLevel
        {
            get
            {
                return WarningLevel == VbaWarningLevel.NotConfigured
                    ? VbaWarningLevel.DisableWithNotification
                    : WarningLevel;
            }
        }

        /// <summary>
        /// True when <paramref name="documentPath"/> sits inside a Trusted Location, which causes
        /// Office to run its macros regardless of the macro setting.
        /// </summary>
        public bool IsInTrustedLocation(string documentPath)
        {
            if (AllTrustedLocationsDisabled || string.IsNullOrEmpty(documentPath))
            {
                return false;
            }

            foreach (TrustedLocation location in TrustedLocations)
            {
                if (location.Contains(documentPath))
                {
                    return true;
                }
            }

            return false;
        }
    }

    /// <summary>A single Trust Center "Trusted Location" entry.</summary>
    public sealed class TrustedLocation
    {
        public TrustedLocation(string path, bool allowSubFolders)
        {
            Path = path;
            AllowSubFolders = allowSubFolders;
        }

        public string Path { get; private set; }

        public bool AllowSubFolders { get; private set; }

        /// <summary>True when <paramref name="documentPath"/> falls under this location.</summary>
        public bool Contains(string documentPath)
        {
            if (string.IsNullOrEmpty(Path) || string.IsNullOrEmpty(documentPath))
            {
                return false;
            }

            string trustedRoot = Normalize(Path);
            string documentFolder = Normalize(GetDirectoryName(documentPath));

            if (string.IsNullOrEmpty(trustedRoot) || string.IsNullOrEmpty(documentFolder))
            {
                return false;
            }

            if (string.Equals(trustedRoot, documentFolder, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return AllowSubFolders
                && documentFolder.StartsWith(trustedRoot + "\\", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Trims a path to a canonical, comparable form: backslash separators, environment
        /// variables expanded, no trailing separator.
        /// </summary>
        private static string Normalize(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return null;
            }

            string expanded = Environment.ExpandEnvironmentVariables(path.Trim());
            expanded = expanded.Replace('/', '\\').TrimEnd('\\');
            return expanded;
        }

        /// <summary>
        /// Directory portion of a path. <see cref="System.IO.Path.GetDirectoryName(string)"/>
        /// throws on the malformed paths that a document's FullName can legitimately hold
        /// (SharePoint URLs, for instance), so this does the split by hand.
        /// </summary>
        private static string GetDirectoryName(string path)
        {
            string normalized = path.Replace('/', '\\');
            int lastSeparator = normalized.LastIndexOf('\\');
            return lastSeparator <= 0 ? null : normalized.Substring(0, lastSeparator);
        }
    }
}
