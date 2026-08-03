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
            // When a trusted publisher is what let the macros run, say so instead of citing the
            // macro setting - the setting is not the reason, and the distinction is the whole
            // point of resolving the signer.
            string opening = status.State == MacroState.RunsSilently && status.Document.IsFromTrustedPublisher
                ? Strings.Get("Detail_RunsSilentlyTrustedPublisher")
                : Strings.Get("Detail_" + status.State);

            var parts = new List<string> { opening };

            if (status.State == MacroState.NoMacros)
            {
                return parts[0];
            }

            // The signature only carries meaning where a VBA project exists; a workbook whose only
            // macros are XLM sheets has nothing to sign.
            if (status.Document.HasVbaProject)
            {
                string signature = DescribeSignature(status.Document, status.State);
                if (signature != null)
                {
                    parts.Add(signature);
                }
            }

            if (status.Document.HasExcel4Macros)
            {
                parts.Add(Strings.Get(status.Document.HasVbaProject ? "Excel4_Also" : "Excel4_Only"));
            }

            return string.Join(" ", parts.ToArray());
        }

        /// <summary>
        /// One sentence about the signature, or null when the state's own wording already covers
        /// it and repeating it would just take up room.
        /// </summary>
        private static string DescribeSignature(DocumentMacroInfo document, MacroState state)
        {
            if (!document.IsVbaSigned)
            {
                // "Unsigned" is already the substance of these two states.
                return state == MacroState.BlockedUnsigned ? null : Strings.Get("Signature_Unsigned");
            }

            switch (document.Signature.Trust)
            {
                case PublisherTrust.Trusted:
                    return Strings.Format("Signature_SignedByTrusted", document.Signature.SignerName);

                case PublisherTrust.NotTrusted:
                    string signer = Strings.Format("Signature_SignedByUntrusted", document.Signature.SignerName);

                    // The state already says it needs trusting; adding why is what turns an
                    // unhelpful "not trusted yet" into something actionable.
                    return state == MacroState.RequiresPublisherTrust && document.Signature.UntrustedReason != null
                        ? signer + " " + Strings.Format("Signature_UntrustedBecause", document.Signature.UntrustedReason)
                        : signer;

                default:
                    // Nothing could be established, so fall back to the bare fact Office reports.
                    return state == MacroState.RequiresPublisherTrust ? null : Strings.Get("Signature_Signed");
            }
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

                    if (status.Document.Signature.Trust == PublisherTrust.Unknown)
                    {
                        // The caveat is only honest while the signature is genuinely unread. Once
                        // the certificate has been resolved, saying we cannot tell would be false.
                        builder.AppendLine(Strings.Get("Signature_Caveat"));
                    }
                    else
                    {
                        builder.AppendLine(Strings.Format("Signature_Thumbprint", status.Document.Signature.Thumbprint));
                    }
                }
            }

            return builder.ToString().TrimEnd();
        }
    }
}
