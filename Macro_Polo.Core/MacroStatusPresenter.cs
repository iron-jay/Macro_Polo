using System;
using System.Collections.Generic;
using System.Text;

namespace Macro_Polo.Core
{
    /// <summary>
    /// Turns a <see cref="MacroStatus"/> into display text and colours.
    /// </summary>
    /// <remarks>
    /// Kept as a lookup over the state rather than as nested conditionals: every state has exactly
    /// one entry, so adding a state is a compile error away from being handled everywhere.
    /// </remarks>
    public static class MacroStatusPresenter
    {
        public static MacroStatusView Describe(MacroStatus status)
        {
            if (status == null) throw new ArgumentNullException("status");

            string title = Strings.Get("Title_" + status.State);
            string detail = BuildDetail(status);
            string toolTip = BuildToolTip(status, title, detail);

            return new MacroStatusView(SeverityFor(status), title, detail, toolTip);
        }

        /// <summary>
        /// Unsigned macros that run with no prompt are the state this add-in exists to surface, so
        /// they get the strongest colour — stronger than a blocked macro, which is inert.
        /// </summary>
        private static Severity SeverityFor(MacroStatus status)
        {
            switch (status.State)
            {
                case MacroState.NoMacros:
                    return Severity.Neutral;

                case MacroState.RunsSilently:
                case MacroState.RunsSilentlyTrustedLocation:
                    return status.Document.IsVbaSigned && !status.Document.HasExcel4Macros
                        ? Severity.Safe
                        : Severity.Danger;

                case MacroState.RequiresUserConsent:
                case MacroState.RequiresPublisherTrust:
                    return Severity.Caution;

                default:
                    return Severity.Information;
            }
        }

        private static string BuildDetail(MacroStatus status)
        {
            var parts = new List<string> { Strings.Get("Detail_" + status.State) };

            if (status.State == MacroState.NoMacros)
            {
                return parts[0];
            }

            // The signature only carries meaning where a VBA project exists; a workbook whose only
            // macros are XLM sheets has nothing to sign.
            if (status.Document.HasVbaProject && !IsSignatureImpliedByState(status.State))
            {
                parts.Add(Strings.Get(status.Document.IsVbaSigned ? "Signature_Signed" : "Signature_Unsigned"));
            }

            if (status.Document.HasExcel4Macros)
            {
                parts.Add(Strings.Get(status.Document.HasVbaProject ? "Excel4_Also" : "Excel4_Only"));
            }

            return string.Join(" ", parts.ToArray());
        }

        /// <summary>
        /// True for states whose own wording already says whether the macros are signed, so the
        /// banner does not repeat it.
        /// </summary>
        private static bool IsSignatureImpliedByState(MacroState state)
        {
            return state == MacroState.RequiresPublisherTrust || state == MacroState.BlockedUnsigned;
        }

        private static string BuildToolTip(MacroStatus status, string title, string detail)
        {
            var builder = new StringBuilder();
            builder.AppendLine(title);
            builder.AppendLine();
            builder.AppendLine(detail);

            if (status.State != MacroState.NoMacros)
            {
                builder.AppendLine();
                builder.AppendLine(Strings.Format("Setting_Effective", Strings.Describe(status.Settings.EffectiveWarningLevel)));

                if (status.Settings.IsManagedByPolicy)
                {
                    builder.AppendLine(Strings.Get("Setting_ManagedByPolicy"));
                }
                else if (status.Settings.WarningLevel == VbaWarningLevel.NotConfigured)
                {
                    builder.AppendLine(Strings.Get("Setting_NotConfigured"));
                }

                if (status.Document.IsVbaSigned)
                {
                    builder.AppendLine();
                    builder.AppendLine(Strings.Get("Signature_Caveat"));
                }
            }

            return builder.ToString().TrimEnd();
        }
    }
}
