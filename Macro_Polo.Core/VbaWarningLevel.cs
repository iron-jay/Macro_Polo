namespace Macro_Polo.Core
{
    /// <summary>
    /// The Trust Center "Macro Settings" value, stored in the registry as
    /// <c>...\Office\&lt;version&gt;\&lt;app&gt;\Security\VBAWarnings</c>.
    /// </summary>
    public enum VbaWarningLevel
    {
        /// <summary>
        /// No value is present in any of the registry hives we consult. Office falls back to
        /// <see cref="DisableWithNotification"/>, so callers should treat this as that value —
        /// it is kept distinct only so the UI can say whether the setting was explicitly configured.
        /// </summary>
        NotConfigured = 0,

        /// <summary>Enable all macros (not recommended).</summary>
        EnableAll = 1,

        /// <summary>Disable all macros with notification. This is the Office default.</summary>
        DisableWithNotification = 2,

        /// <summary>Disable all macros except digitally signed macros.</summary>
        DisableExceptSigned = 3,

        /// <summary>Disable all macros without notification.</summary>
        DisableAll = 4
    }
}
