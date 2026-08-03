using System;
using System.Collections.Generic;
using System.Globalization;

namespace Macro_Polo.Core
{
    /// <summary>
    /// Resolves the effective Trust Center settings for one Office application.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Precedence matters, and it is the opposite of the obvious reading order. Office applies
    /// Group Policy over the user's own preference, so the policy hives have to be consulted
    /// first: a managed machine normally still has a stale value sitting in the user preference
    /// key, and trusting that value reports the wrong answer on exactly the machines this add-in
    /// exists to serve.
    /// </para>
    /// <para>Order, highest priority first:</para>
    /// <list type="number">
    ///   <item><description>HKLM\Software\Policies\Microsoft\Office\{version}\{app}\Security</description></item>
    ///   <item><description>HKCU\Software\Policies\Microsoft\Office\{version}\{app}\Security</description></item>
    ///   <item><description>HKCU\Software\Microsoft\Office\{version}\{app}\Security</description></item>
    /// </list>
    /// </remarks>
    public sealed class OfficeSecurityReader
    {
        private const string UserPreferenceRoot = @"Software\Microsoft\Office\{0}\{1}\Security";
        private const string PolicyRoot = @"Software\Policies\Microsoft\Office\{0}\{1}\Security";

        private readonly IRegistryValueSource _registry;
        private readonly string _policyPath;
        private readonly string _preferencePath;

        /// <param name="registry">Registry access.</param>
        /// <param name="officeVersion">Major Office version, for example "16.0".</param>
        /// <param name="applicationName">Registry name of the host application: "Word" or "Excel".</param>
        public OfficeSecurityReader(IRegistryValueSource registry, string officeVersion, string applicationName)
        {
            if (registry == null) throw new ArgumentNullException("registry");
            if (string.IsNullOrEmpty(officeVersion)) throw new ArgumentException("Office version is required.", "officeVersion");
            if (string.IsNullOrEmpty(applicationName)) throw new ArgumentException("Application name is required.", "applicationName");

            _registry = registry;
            _policyPath = string.Format(CultureInfo.InvariantCulture, PolicyRoot, officeVersion, applicationName);
            _preferencePath = string.Format(CultureInfo.InvariantCulture, UserPreferenceRoot, officeVersion, applicationName);
        }

        public MacroSecuritySettings Read()
        {
            var settings = new MacroSecuritySettings();

            bool fromPolicy;
            int? warnings = ReadWarningLevel(out fromPolicy);

            settings.WarningLevel = ToWarningLevel(warnings);
            settings.IsManagedByPolicy = warnings.HasValue && fromPolicy;
            settings.BlockMacrosFromInternet = ReadFlag("blockcontentexecutionfrominternet");
            settings.AllTrustedLocationsDisabled = ReadTrustedLocationFlag("AllLocationsDisabled");

            settings.TrustedLocations.AddRange(ReadTrustedLocations(RegistryRoot.LocalMachine, _policyPath));
            settings.TrustedLocations.AddRange(ReadTrustedLocations(RegistryRoot.CurrentUser, _policyPath));
            settings.TrustedLocations.AddRange(ReadTrustedLocations(RegistryRoot.CurrentUser, _preferencePath));

            settings.AllTrustedDocumentsDisabled = ReadTrustedDocumentFlag("DisableTrustedDocuments");
            settings.TrustedDocuments.AddRange(ReadTrustedDocuments());

            return settings;
        }

        /// <summary>
        /// Reads VBAWarnings from the highest-priority hive that defines it.
        /// </summary>
        private int? ReadWarningLevel(out bool fromPolicy)
        {
            fromPolicy = true;

            int? value = ReadInt(RegistryRoot.LocalMachine, _policyPath, "VBAWarnings");
            if (value.HasValue) return value;

            value = ReadInt(RegistryRoot.CurrentUser, _policyPath, "VBAWarnings");
            if (value.HasValue) return value;

            fromPolicy = false;
            return ReadInt(RegistryRoot.CurrentUser, _preferencePath, "VBAWarnings");
        }

        /// <summary>
        /// True when any hive sets the named DWORD to a non-zero value. Policy hives win, but for
        /// a boolean hardening flag "set anywhere" and "set by the winning hive" only differ if a
        /// policy explicitly clears it, which is handled by the ordering below.
        /// </summary>
        private bool ReadFlag(string valueName)
        {
            int? value = ReadInt(RegistryRoot.LocalMachine, _policyPath, valueName)
                ?? ReadInt(RegistryRoot.CurrentUser, _policyPath, valueName)
                ?? ReadInt(RegistryRoot.CurrentUser, _preferencePath, valueName);

            return value.GetValueOrDefault() != 0;
        }

        private bool ReadTrustedLocationFlag(string valueName)
        {
            int? value = ReadInt(RegistryRoot.LocalMachine, _policyPath + @"\Trusted Locations", valueName)
                ?? ReadInt(RegistryRoot.CurrentUser, _policyPath + @"\Trusted Locations", valueName)
                ?? ReadInt(RegistryRoot.CurrentUser, _preferencePath + @"\Trusted Locations", valueName);

            return value.GetValueOrDefault() != 0;
        }

        private bool ReadTrustedDocumentFlag(string valueName)
        {
            int? value = ReadInt(RegistryRoot.LocalMachine, _policyPath + @"\Trusted Documents", valueName)
                ?? ReadInt(RegistryRoot.CurrentUser, _policyPath + @"\Trusted Documents", valueName)
                ?? ReadInt(RegistryRoot.CurrentUser, _preferencePath + @"\Trusted Documents", valueName);

            return value.GetValueOrDefault() != 0;
        }

        /// <summary>
        /// Reads the trust records. Each document is a value whose <em>name</em> is the path; the
        /// data is a timestamp and flags whose meaning is not documented, and is deliberately not
        /// interpreted here - the presence of the record is what Office acts on.
        /// </summary>
        private IEnumerable<string> ReadTrustedDocuments()
        {
            string path = _preferencePath + @"\Trusted Documents\TrustRecords";
            var documents = new List<string>();

            foreach (string valueName in _registry.GetValueNames(RegistryRoot.CurrentUser, path))
            {
                string normalized = MacroSecuritySettings.NormalizePath(valueName);
                if (!string.IsNullOrEmpty(normalized))
                {
                    documents.Add(normalized);
                }
            }

            return documents;
        }

        private IEnumerable<TrustedLocation> ReadTrustedLocations(RegistryRoot root, string securityPath)
        {
            string locationsPath = securityPath + @"\Trusted Locations";
            var locations = new List<TrustedLocation>();

            foreach (string subKeyName in _registry.GetSubKeyNames(root, locationsPath))
            {
                string keyPath = locationsPath + "\\" + subKeyName;
                var path = _registry.GetValue(root, keyPath, "Path") as string;

                if (string.IsNullOrEmpty(path))
                {
                    continue;
                }

                bool allowSubFolders = ReadInt(root, keyPath, "AllowSubFolders").GetValueOrDefault() != 0;
                locations.Add(new TrustedLocation(path, allowSubFolders));
            }

            return locations;
        }

        /// <summary>
        /// Reads a value that should be a DWORD. Hand-edited and script-deployed keys are
        /// routinely written as strings, so anything convertible is accepted and anything else is
        /// treated as absent rather than throwing.
        /// </summary>
        private int? ReadInt(RegistryRoot root, string subKeyPath, string valueName)
        {
            object raw = _registry.GetValue(root, subKeyPath, valueName);
            if (raw == null)
            {
                return null;
            }

            try
            {
                return Convert.ToInt32(raw, CultureInfo.InvariantCulture);
            }
            catch (Exception ex)
            {
                Log.Warn("Unexpected type for " + subKeyPath + "\\" + valueName + ": " + raw.GetType().Name, ex);
                return null;
            }
        }

        private static VbaWarningLevel ToWarningLevel(int? value)
        {
            if (!value.HasValue)
            {
                return VbaWarningLevel.NotConfigured;
            }

            return Enum.IsDefined(typeof(VbaWarningLevel), value.Value)
                ? (VbaWarningLevel)value.Value
                : VbaWarningLevel.NotConfigured;
        }
    }
}
