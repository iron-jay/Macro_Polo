using Macro_Polo.Core;
using Xunit;

namespace Macro_Polo.Core.Tests
{
    public class PresentationTests
    {
        /// <summary>
        /// The state the add-in exists to catch: macros that will run with no prompt and no
        /// signature. The original palette gave this the same amber as a macro the user still had
        /// to approve.
        /// </summary>
        [Fact]
        public void Unsigned_macros_that_run_unprompted_get_the_strongest_severity()
        {
            MacroStatusView view = Describe(signed: false, level: VbaWarningLevel.EnableAll);

            Assert.Equal(Severity.Danger, view.Severity);
        }

        [Fact]
        public void Signed_macros_that_run_unprompted_are_shown_as_safe()
        {
            MacroStatusView view = Describe(signed: true, level: VbaWarningLevel.EnableAll);

            Assert.Equal(Severity.Safe, view.Severity);
        }

        [Fact]
        public void Blocked_macros_are_informational_rather_than_alarming()
        {
            MacroStatusView view = Describe(signed: false, level: VbaWarningLevel.DisableAll);

            Assert.Equal(Severity.Information, view.Severity);
        }

        [Fact]
        public void A_document_without_macros_is_neutral()
        {
            var status = MacroStatusEvaluator.Evaluate(
                new DocumentMacroInfo { HasVbaProject = false },
                new MacroSecuritySettings());

            Assert.Equal(Severity.Neutral, MacroStatusPresenter.Describe(status).Severity);
        }

        /// <summary>
        /// Severity is also carried by a glyph, so the banner does not rely on colour alone.
        /// </summary>
        [Fact]
        public void Each_severity_has_its_own_glyph()
        {
            Assert.NotEqual(
                Describe(signed: false, level: VbaWarningLevel.EnableAll).Glyph,
                Describe(signed: true, level: VbaWarningLevel.EnableAll).Glyph);
        }

        [Fact]
        public void Every_state_produces_text_rather_than_a_missing_resource_name()
        {
            foreach (MacroState state in System.Enum.GetValues(typeof(MacroState)))
            {
                string title = Strings.Get("Title_" + state);
                string detail = Strings.Get("Detail_" + state);

                Assert.False(title.StartsWith("Title_"), "No title resource for " + state);
                Assert.False(detail.StartsWith("Detail_"), "No detail resource for " + state);
            }
        }

        [Fact]
        public void The_tooltip_explains_what_the_signature_flag_does_not_prove()
        {
            MacroStatusView view = Describe(signed: true, level: VbaWarningLevel.DisableExceptSigned);

            Assert.Contains(Strings.Get("Signature_Caveat"), view.ToolTip);
        }

        [Fact]
        public void The_tooltip_says_when_the_setting_is_imposed_by_policy()
        {
            var settings = new MacroSecuritySettings
            {
                WarningLevel = VbaWarningLevel.DisableAll,
                IsManagedByPolicy = true
            };

            var status = MacroStatusEvaluator.Evaluate(
                new DocumentMacroInfo { HasVbaProject = true },
                settings);

            Assert.Contains(Strings.Get("Setting_ManagedByPolicy"), MacroStatusPresenter.Describe(status).ToolTip);
        }

        [Fact]
        public void Excel4_macro_sheets_are_called_out_in_the_detail_text()
        {
            var document = new DocumentMacroInfo { HasVbaProject = false, HasExcel4Macros = true };
            var status = MacroStatusEvaluator.Evaluate(document, new MacroSecuritySettings());

            Assert.Contains(Strings.Get("Excel4_Only"), MacroStatusPresenter.Describe(status).Detail);
        }

        private static MacroStatusView Describe(bool signed, VbaWarningLevel level)
        {
            var status = MacroStatusEvaluator.Evaluate(
                new DocumentMacroInfo { HasVbaProject = true, IsVbaSigned = signed },
                new MacroSecuritySettings { WarningLevel = level });

            return MacroStatusPresenter.Describe(status);
        }
    }
}
