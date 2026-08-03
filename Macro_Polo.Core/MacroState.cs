namespace Macro_Polo.Core
{
    /// <summary>
    /// What will actually happen to the macros in a document, given the document's own
    /// properties and the effective Trust Center configuration.
    /// </summary>
    public enum MacroState
    {
        /// <summary>The document contains no macro code at all.</summary>
        NoMacros,

        /// <summary>
        /// Macros run as soon as the document opens, with no prompt, because the file lives in a
        /// Trusted Location. Trusted Locations bypass the macro settings entirely.
        /// </summary>
        RunsSilentlyTrustedLocation,

        /// <summary>Macros run with no prompt because macro settings are set to "enable all".</summary>
        RunsSilently,

        /// <summary>
        /// Macros are blocked until the user clicks "Enable Content" on the trust bar.
        /// </summary>
        RequiresUserConsent,

        /// <summary>
        /// The macros are signed and macro settings allow signed macros, but they will only run
        /// once the signing certificate is added to Trusted Publishers.
        /// </summary>
        RequiresPublisherTrust,

        /// <summary>
        /// Macro settings allow signed macros only, and these macros are unsigned. There is no
        /// prompt and no way for the user to run them.
        /// </summary>
        BlockedUnsigned,

        /// <summary>Macro settings disable all macros without notification.</summary>
        BlockedByPolicy,

        /// <summary>
        /// The file carries a mark of the web and the "block macros from running in Office files
        /// from the internet" policy is enabled. This cannot be overridden from inside Office.
        /// </summary>
        BlockedFromInternet
    }
}
