using Macro_Polo.Core;
using Xunit;

namespace Macro_Polo.Core.Tests
{
    /// <summary>
    /// The decision table. These are the cases the original two-branch implementation got wrong:
    /// it treated only "disable all without notification" as blocking, so a document that Office
    /// would refuse to run was reported as runnable.
    /// </summary>
    public class MacroStatusEvaluatorTests
    {
        [Fact]
        public void A_document_without_macros_is_reported_as_such_whatever_the_settings_say()
        {
            MacroStatus status = Evaluate(
                Document(hasVba: false),
                Settings(VbaWarningLevel.EnableAll));

            Assert.Equal(MacroState.NoMacros, status.State);
        }

        [Theory]
        [InlineData(VbaWarningLevel.EnableAll, MacroState.RunsSilently)]
        [InlineData(VbaWarningLevel.DisableWithNotification, MacroState.RequiresUserConsent)]
        [InlineData(VbaWarningLevel.DisableExceptSigned, MacroState.BlockedUnsigned)]
        [InlineData(VbaWarningLevel.DisableAll, MacroState.BlockedByPolicy)]
        public void Unsigned_macros_follow_the_macro_setting(VbaWarningLevel level, MacroState expected)
        {
            MacroStatus status = Evaluate(Document(signed: false), Settings(level));

            Assert.Equal(expected, status.State);
        }

        [Theory]
        [InlineData(VbaWarningLevel.EnableAll, MacroState.RunsSilently)]
        [InlineData(VbaWarningLevel.DisableWithNotification, MacroState.RequiresUserConsent)]
        [InlineData(VbaWarningLevel.DisableExceptSigned, MacroState.RequiresPublisherTrust)]
        [InlineData(VbaWarningLevel.DisableAll, MacroState.BlockedByPolicy)]
        public void Signed_macros_follow_the_macro_setting(VbaWarningLevel level, MacroState expected)
        {
            MacroStatus status = Evaluate(Document(signed: true), Settings(level));

            Assert.Equal(expected, status.State);
        }

        /// <summary>
        /// The specific false reassurance in the original: at level 3 an unsigned macro cannot run
        /// at all, but it was shown with the same "contains an unsigned macro" wording used when
        /// macros were runnable.
        /// </summary>
        [Fact]
        public void Unsigned_macros_are_blocked_when_only_signed_macros_are_permitted()
        {
            MacroStatus status = Evaluate(
                Document(signed: false),
                Settings(VbaWarningLevel.DisableExceptSigned));

            Assert.Equal(MacroState.BlockedUnsigned, status.State);
            Assert.True(status.IsBlocked);
            Assert.False(status.RunsWithoutPrompting);
        }

        [Fact]
        public void An_unset_macro_setting_is_treated_as_the_office_default()
        {
            MacroStatus status = Evaluate(Document(), Settings(VbaWarningLevel.NotConfigured));

            Assert.Equal(MacroState.RequiresUserConsent, status.State);
        }

        [Fact]
        public void A_trusted_location_runs_macros_regardless_of_the_macro_setting()
        {
            MacroSecuritySettings settings = Settings(VbaWarningLevel.DisableAll);
            settings.TrustedLocations.Add(new TrustedLocation(@"C:\Templates", allowSubFolders: false));

            MacroStatus status = Evaluate(Document(path: @"C:\Templates\budget.xlsm"), settings);

            Assert.Equal(MacroState.RunsSilentlyTrustedLocation, status.State);
            Assert.True(status.RunsWithoutPrompting);
        }

        [Fact]
        public void A_trusted_location_only_covers_subfolders_when_it_says_so()
        {
            MacroSecuritySettings settings = Settings(VbaWarningLevel.DisableAll);
            settings.TrustedLocations.Add(new TrustedLocation(@"C:\Templates", allowSubFolders: false));

            MacroStatus status = Evaluate(Document(path: @"C:\Templates\Team\budget.xlsm"), settings);

            Assert.Equal(MacroState.BlockedByPolicy, status.State);
        }

        [Fact]
        public void A_trusted_location_covers_subfolders_when_allowed()
        {
            MacroSecuritySettings settings = Settings(VbaWarningLevel.DisableAll);
            settings.TrustedLocations.Add(new TrustedLocation(@"C:\Templates\", allowSubFolders: true));

            MacroStatus status = Evaluate(Document(path: @"C:\Templates\Team\budget.xlsm"), settings);

            Assert.Equal(MacroState.RunsSilentlyTrustedLocation, status.State);
        }

        /// <summary>A folder that merely starts with the same characters is not inside it.</summary>
        [Fact]
        public void A_trusted_location_does_not_match_a_sibling_with_a_shared_prefix()
        {
            MacroSecuritySettings settings = Settings(VbaWarningLevel.DisableAll);
            settings.TrustedLocations.Add(new TrustedLocation(@"C:\Templates", allowSubFolders: true));

            MacroStatus status = Evaluate(Document(path: @"C:\Templates_Old\budget.xlsm"), settings);

            Assert.Equal(MacroState.BlockedByPolicy, status.State);
        }

        [Fact]
        public void Disabling_all_trusted_locations_takes_them_out_of_the_reckoning()
        {
            MacroSecuritySettings settings = Settings(VbaWarningLevel.DisableAll);
            settings.AllTrustedLocationsDisabled = true;
            settings.TrustedLocations.Add(new TrustedLocation(@"C:\Templates", allowSubFolders: true));

            MacroStatus status = Evaluate(Document(path: @"C:\Templates\budget.xlsm"), settings);

            Assert.Equal(MacroState.BlockedByPolicy, status.State);
        }

        [Fact]
        public void A_marked_file_is_blocked_when_the_internet_policy_is_on()
        {
            MacroSecuritySettings settings = Settings(VbaWarningLevel.EnableAll);
            settings.BlockMacrosFromInternet = true;

            MacroStatus status = Evaluate(Document(motw: true), settings);

            Assert.Equal(MacroState.BlockedFromInternet, status.State);
        }

        [Fact]
        public void The_internet_policy_only_applies_to_marked_files()
        {
            MacroSecuritySettings settings = Settings(VbaWarningLevel.EnableAll);
            settings.BlockMacrosFromInternet = true;

            MacroStatus status = Evaluate(Document(motw: false), settings);

            Assert.Equal(MacroState.RunsSilently, status.State);
        }

        /// <summary>A Trusted Location exempts a file from the internet block, as it does in Office.</summary>
        [Fact]
        public void A_trusted_location_wins_over_the_internet_policy()
        {
            MacroSecuritySettings settings = Settings(VbaWarningLevel.DisableAll);
            settings.BlockMacrosFromInternet = true;
            settings.TrustedLocations.Add(new TrustedLocation(@"C:\Templates", allowSubFolders: false));

            MacroStatus status = Evaluate(Document(path: @"C:\Templates\budget.xlsm", motw: true), settings);

            Assert.Equal(MacroState.RunsSilentlyTrustedLocation, status.State);
        }

        [Fact]
        public void Excel4_macro_sheets_count_as_macros_even_without_a_vba_project()
        {
            var document = new DocumentMacroInfo { HasVbaProject = false, HasExcel4Macros = true };

            MacroStatus status = Evaluate(document, Settings(VbaWarningLevel.DisableWithNotification));

            Assert.True(document.HasMacros);
            Assert.Equal(MacroState.RequiresUserConsent, status.State);
        }

        /// <summary>
        /// XLM sheets cannot carry a VBA signature, so "signed macros only" blocks them even in a
        /// workbook whose VBA project is signed.
        /// </summary>
        [Fact]
        public void Excel4_macro_sheets_are_blocked_when_only_signed_macros_are_permitted()
        {
            var document = new DocumentMacroInfo { HasVbaProject = false, HasExcel4Macros = true };

            MacroStatus status = Evaluate(document, Settings(VbaWarningLevel.DisableExceptSigned));

            Assert.Equal(MacroState.BlockedUnsigned, status.State);
        }

        private static MacroStatus Evaluate(DocumentMacroInfo document, MacroSecuritySettings settings)
        {
            return MacroStatusEvaluator.Evaluate(document, settings);
        }

        private static DocumentMacroInfo Document(
            bool hasVba = true,
            bool signed = false,
            string path = @"C:\Users\someone\Documents\report.docm",
            bool motw = false)
        {
            return new DocumentMacroInfo
            {
                HasVbaProject = hasVba,
                IsVbaSigned = signed,
                FullPath = path,
                HasMarkOfTheWeb = motw
            };
        }

        private static MacroSecuritySettings Settings(VbaWarningLevel level)
        {
            return new MacroSecuritySettings { WarningLevel = level };
        }
    }
}
