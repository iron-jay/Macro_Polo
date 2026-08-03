using System;
using System.Globalization;
using System.IO;

namespace Macro_Polo.Core
{
    /// <summary>
    /// Detects the mark of the web: the <c>Zone.Identifier</c> alternate data stream that Windows
    /// attaches to files downloaded from the internet.
    /// </summary>
    /// <remarks>
    /// This matters because the "block macros from running in Office files from the Internet"
    /// policy keys off it, and that block cannot be lifted from inside Office. A document that is
    /// marked and blocked is in a genuinely different state from one that merely needs
    /// "Enable Content" clicked.
    /// </remarks>
    public static class MarkOfTheWeb
    {
        /// <summary>Internet zone.</summary>
        private const int ZoneInternet = 3;

        /// <summary>Restricted sites zone.</summary>
        private const int ZoneRestricted = 4;

        /// <summary>
        /// True when the file at <paramref name="fullPath"/> is marked as coming from the internet
        /// or a restricted site. Returns false for anything not on a local or UNC path — a
        /// document opened straight from SharePoint or a URL has no local stream to inspect.
        /// </summary>
        public static bool IsPresent(string fullPath)
        {
            if (string.IsNullOrEmpty(fullPath) || fullPath.IndexOf("://", StringComparison.Ordinal) >= 0)
            {
                return false;
            }

            try
            {
                if (!File.Exists(fullPath))
                {
                    return false;
                }

                // Alternate data streams are addressed with a "file:stream" path. The framework
                // file APIs pass this through to the Win32 layer unchanged.
                string streamPath = fullPath + ":Zone.Identifier";

                using (var reader = new StreamReader(streamPath))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        int zoneId;
                        if (TryParseZoneId(line, out zoneId))
                        {
                            return zoneId == ZoneInternet || zoneId == ZoneRestricted;
                        }
                    }
                }
            }
            catch (FileNotFoundException)
            {
                // No Zone.Identifier stream: the ordinary case for a local file.
            }
            catch (Exception ex)
            {
                // A volume without ADS support (FAT32, some network shares) throws various things.
                Log.Warn("Could not read the zone identifier for " + fullPath, ex);
            }

            return false;
        }

        private static bool TryParseZoneId(string line, out int zoneId)
        {
            zoneId = 0;

            if (line == null)
            {
                return false;
            }

            string trimmed = line.Trim();
            const string Prefix = "ZoneId=";

            if (!trimmed.StartsWith(Prefix, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return int.TryParse(
                trimmed.Substring(Prefix.Length).Trim(),
                NumberStyles.Integer,
                CultureInfo.InvariantCulture,
                out zoneId);
        }
    }
}
