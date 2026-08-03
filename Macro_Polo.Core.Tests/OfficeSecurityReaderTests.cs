using System;
using Macro_Polo.Core;
using Xunit;

namespace Macro_Polo.Core.Tests
{
    /// <summary>
    /// Hive precedence. The original implementation read the user preference first and only
    /// consulted Group Policy when that was absent, which is backwards: Office applies policy over
    /// preference. On a managed machine the stale preference value is usually still present, so the
    /// wrong hive won on exactly the machines this add-in is meant for.
    /// </summary>
    public class OfficeSecurityReaderTests
    {
        private const string Policy = @"Software\Policies\Microsoft\Office\16.0\Word\Security";
        private const string Preference = @"Software\Microsoft\Office\16.0\Word\Security";

        [Fact]
        public void Machine_policy_beats_user_policy_and_user_preference()
        {
            var registry = new FakeRegistry()
                .Set(RegistryRoot.LocalMachine, Policy, "VBAWarnings", 4)
                .Set(RegistryRoot.CurrentUser, Policy, "VBAWarnings", 3)
                .Set(RegistryRoot.CurrentUser, Preference, "VBAWarnings", 1);

            MacroSecuritySettings settings = Read(registry);

            Assert.Equal(VbaWarningLevel.DisableAll, settings.WarningLevel);
            Assert.True(settings.IsManagedByPolicy);
        }

        [Fact]
        public void User_policy_beats_user_preference()
        {
            var registry = new FakeRegistry()
                .Set(RegistryRoot.CurrentUser, Policy, "VBAWarnings", 3)
                .Set(RegistryRoot.CurrentUser, Preference, "VBAWarnings", 1);

            MacroSecuritySettings settings = Read(registry);

            Assert.Equal(VbaWarningLevel.DisableExceptSigned, settings.WarningLevel);
            Assert.True(settings.IsManagedByPolicy);
        }

        [Fact]
        public void The_user_preference_is_used_when_no_policy_is_set()
        {
            var registry = new FakeRegistry()
                .Set(RegistryRoot.CurrentUser, Preference, "VBAWarnings", 1);

            MacroSecuritySettings settings = Read(registry);

            Assert.Equal(VbaWarningLevel.EnableAll, settings.WarningLevel);
            Assert.False(settings.IsManagedByPolicy);
        }

        [Fact]
        public void An_absent_setting_reports_as_not_configured_but_behaves_as_the_default()
        {
            MacroSecuritySettings settings = Read(new FakeRegistry());

            Assert.Equal(VbaWarningLevel.NotConfigured, settings.WarningLevel);
            Assert.Equal(VbaWarningLevel.DisableWithNotification, settings.EffectiveWarningLevel);
            Assert.False(settings.IsManagedByPolicy);
        }

        /// <summary>
        /// Deployment scripts routinely write these values as strings. The original code cast the
        /// boxed value straight to int, which threw.
        /// </summary>
        [Fact]
        public void A_setting_stored_as_a_string_is_still_understood()
        {
            var registry = new FakeRegistry()
                .Set(RegistryRoot.CurrentUser, Preference, "VBAWarnings", "4");

            Assert.Equal(VbaWarningLevel.DisableAll, Read(registry).WarningLevel);
        }

        [Fact]
        public void A_nonsense_setting_falls_back_to_the_default_rather_than_throwing()
        {
            var registry = new FakeRegistry()
                .Set(RegistryRoot.CurrentUser, Preference, "VBAWarnings", "not a number");

            MacroSecuritySettings settings = Read(registry);

            Assert.Equal(VbaWarningLevel.NotConfigured, settings.WarningLevel);
            Assert.Equal(VbaWarningLevel.DisableWithNotification, settings.EffectiveWarningLevel);
        }

        [Fact]
        public void An_out_of_range_setting_falls_back_to_the_default()
        {
            var registry = new FakeRegistry()
                .Set(RegistryRoot.CurrentUser, Preference, "VBAWarnings", 99);

            Assert.Equal(VbaWarningLevel.NotConfigured, Read(registry).WarningLevel);
        }

        [Fact]
        public void The_internet_block_policy_is_read()
        {
            var registry = new FakeRegistry()
                .Set(RegistryRoot.LocalMachine, Policy, "blockcontentexecutionfrominternet", 1);

            Assert.True(Read(registry).BlockMacrosFromInternet);
        }

        [Fact]
        public void Trusted_locations_are_collected_from_every_hive()
        {
            var registry = new FakeRegistry()
                .Set(RegistryRoot.CurrentUser, Preference + @"\Trusted Locations\Location0", "Path", @"C:\Templates")
                .Set(RegistryRoot.CurrentUser, Preference + @"\Trusted Locations\Location0", "AllowSubFolders", 1)
                .Set(RegistryRoot.LocalMachine, Policy + @"\Trusted Locations\Location1", "Path", @"D:\Shared");

            MacroSecuritySettings settings = Read(registry);

            Assert.Equal(2, settings.TrustedLocations.Count);
            Assert.True(settings.IsInTrustedLocation(@"C:\Templates\Team\budget.xlsm"));
            Assert.True(settings.IsInTrustedLocation(@"D:\Shared\report.docm"));
            Assert.False(settings.IsInTrustedLocation(@"D:\Shared\Sub\report.docm"));
        }

        [Fact]
        public void A_trusted_location_entry_without_a_path_is_ignored()
        {
            var registry = new FakeRegistry()
                .Set(RegistryRoot.CurrentUser, Preference + @"\Trusted Locations\Location0", "AllowSubFolders", 1);

            Assert.Empty(Read(registry).TrustedLocations);
        }

        [Fact]
        public void All_locations_disabled_is_read()
        {
            var registry = new FakeRegistry()
                .Set(RegistryRoot.CurrentUser, Preference + @"\Trusted Locations", "AllLocationsDisabled", 1)
                .Set(RegistryRoot.CurrentUser, Preference + @"\Trusted Locations\Location0", "Path", @"C:\Templates");

            MacroSecuritySettings settings = Read(registry);

            Assert.True(settings.AllTrustedLocationsDisabled);
            Assert.False(settings.IsInTrustedLocation(@"C:\Templates\budget.xlsm"));
        }

        /// <summary>
        /// The registry path is built from the host's version. Hard-coding 16.0, as the original
        /// did, silently reads nothing at all on any other release of Office.
        /// </summary>
        [Fact]
        public void The_office_version_selects_the_registry_path()
        {
            var registry = new FakeRegistry()
                .Set(RegistryRoot.CurrentUser, @"Software\Microsoft\Office\15.0\Word\Security", "VBAWarnings", 4);

            var reader = new OfficeSecurityReader(registry, "15.0", "Word");

            Assert.Equal(VbaWarningLevel.DisableAll, reader.Read().WarningLevel);
            Assert.Equal(VbaWarningLevel.NotConfigured, Read(registry).WarningLevel);
        }

        [Fact]
        public void Word_and_excel_settings_are_kept_apart()
        {
            var registry = new FakeRegistry()
                .Set(RegistryRoot.CurrentUser, @"Software\Microsoft\Office\16.0\Excel\Security", "VBAWarnings", 4);

            Assert.Equal(VbaWarningLevel.NotConfigured, Read(registry).WarningLevel);
            Assert.Equal(
                VbaWarningLevel.DisableAll,
                new OfficeSecurityReader(registry, "16.0", "Excel").Read().WarningLevel);
        }

        /// <summary>
        /// Trust records are stored one value per document, with the path as the value name. The
        /// data is a timestamp and flags whose meaning is undocumented, so only the presence of
        /// the record is used.
        /// </summary>
        [Fact]
        public void Trusted_documents_are_read_from_the_trust_records()
        {
            var registry = new FakeRegistry()
                .Set(RegistryRoot.CurrentUser, Preference + @"\Trusted Documents\TrustRecords",
                     @"%USERPROFILE%/Downloads/quarterly.docm", new byte[] { 1, 2, 3, 4 });

            MacroSecuritySettings settings = Read(registry);

            string expected = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), @"Downloads\quarterly.docm");

            Assert.Single(settings.TrustedDocuments);
            Assert.True(settings.IsTrustedDocument(expected));
            Assert.False(settings.IsTrustedDocument(@"C:\elsewhere\quarterly.docm"));
        }

        [Fact]
        public void Turning_off_trusted_documents_is_read()
        {
            var registry = new FakeRegistry()
                .Set(RegistryRoot.CurrentUser, Preference + @"\Trusted Documents", "DisableTrustedDocuments", 1)
                .Set(RegistryRoot.CurrentUser, Preference + @"\Trusted Documents\TrustRecords", @"C:\x\y.docm", new byte[] { 1 });

            MacroSecuritySettings settings = Read(registry);

            Assert.True(settings.AllTrustedDocumentsDisabled);
            Assert.False(settings.IsTrustedDocument(@"C:\x\y.docm"));
        }

        [Fact]
        public void Turning_off_trusted_documents_by_policy_is_read()
        {
            var registry = new FakeRegistry()
                .Set(RegistryRoot.LocalMachine, Policy + @"\Trusted Documents", "DisableTrustedDocuments", 1);

            Assert.True(Read(registry).AllTrustedDocumentsDisabled);
        }

        [Fact]
        public void No_trust_records_means_no_trusted_documents()
        {
            Assert.Empty(Read(new FakeRegistry()).TrustedDocuments);
        }

        private static MacroSecuritySettings Read(FakeRegistry registry)
        {
            return new OfficeSecurityReader(registry, "16.0", "Word").Read();
        }
    }
}

