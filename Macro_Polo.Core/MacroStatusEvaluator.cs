using System;

namespace Macro_Polo.Core
{
    /// <summary>
    /// Works out what Office will do with a document's macros. This is a pure function of its
    /// inputs so that the decision table can be tested without Office or the registry.
    /// </summary>
    public static class MacroStatusEvaluator
    {
        /// <summary>
        /// Evaluates <paramref name="document"/> against <paramref name="settings"/>.
        /// </summary>
        /// <remarks>
        /// The order of the checks mirrors the order Office itself applies them, which is why a
        /// Trusted Location short-circuits everything below it: a file in a Trusted Location runs
        /// its macros even when the macro setting is "disable all without notification", and even
        /// when it came from the internet.
        /// </remarks>
        public static MacroStatus Evaluate(DocumentMacroInfo document, MacroSecuritySettings settings)
        {
            if (document == null) throw new ArgumentNullException("document");
            if (settings == null) throw new ArgumentNullException("settings");

            if (!document.HasMacros)
            {
                return new MacroStatus(MacroState.NoMacros, document, settings);
            }

            if (settings.IsInTrustedLocation(document.FullPath))
            {
                return new MacroStatus(MacroState.RunsSilentlyTrustedLocation, document, settings);
            }

            if (settings.BlockMacrosFromInternet && document.HasMarkOfTheWeb)
            {
                return new MacroStatus(MacroState.BlockedFromInternet, document, settings);
            }

            // "Disable all without notification" is the one setting a remembered trust decision
            // does not survive, so it is checked first.
            if (settings.EffectiveWarningLevel != VbaWarningLevel.DisableAll
                && settings.IsTrustedDocument(document.FullPath))
            {
                return new MacroStatus(MacroState.RunsSilentlyTrustedDocument, document, settings);
            }

            MacroState state;
            switch (settings.EffectiveWarningLevel)
            {
                case VbaWarningLevel.EnableAll:
                    state = MacroState.RunsSilently;
                    break;

                case VbaWarningLevel.DisableWithNotification:
                    // A trusted publisher is exactly what this setting defers to: the trust bar
                    // never appears and the macros have already run by the time anyone looks.
                    // Reporting these as "blocked until you allow them" understated what happened.
                    state = document.IsFromTrustedPublisher
                        ? MacroState.RunsSilently
                        : MacroState.RequiresUserConsent;
                    break;

                case VbaWarningLevel.DisableExceptSigned:
                    // Excel 4.0 macro sheets cannot carry a VBA signature, so under this setting
                    // they are blocked outright even in a workbook whose VBA project is signed.
                    if (document.IsVbaSigned && document.HasVbaProject)
                    {
                        state = document.Signature.Trust == PublisherTrust.Trusted
                            ? MacroState.RunsSilently
                            : MacroState.RequiresPublisherTrust;
                    }
                    else
                    {
                        state = MacroState.BlockedUnsigned;
                    }
                    break;

                case VbaWarningLevel.DisableAll:
                    state = MacroState.BlockedByPolicy;
                    break;

                default:
                    // EffectiveWarningLevel never returns NotConfigured, but an out-of-range value
                    // read from the registry lands here. Office treats anything it does not
                    // recognise as the default, so we do too.
                    state = MacroState.RequiresUserConsent;
                    break;
            }

            return new MacroStatus(state, document, settings);
        }
    }
}
