namespace Macro_Polo.Core
{
    /// <summary>
    /// Where the add-in's own settings live, and the order they win in.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Group Policy first, then a machine-wide default, then the user's own preference - the same
    /// shape Office uses for its macro settings, so an administrator can impose a value and a user
    /// can choose one where none is imposed.
    /// </para>
    /// <para>
    /// The policy key matches the ADMX shipped in the installer's <c>policies</c> folder. Changing
    /// either without the other silently detaches the templates from the add-in, so they are
    /// defined here once and referenced from both.
    /// </para>
    /// </remarks>
    internal static class OptionScopes
    {
        /// <summary>Set by Group Policy. Overrides everything below it.</summary>
        internal const string Policy = @"Software\Policies\iron-jay\Macro Polo";

        /// <summary>Machine-wide default, which the installer can write with AUTOSHOW=n.</summary>
        internal const string Machine = @"Software\iron-jay\Macro Polo";

        /// <summary>The user's own preference, used when nothing above is set.</summary>
        internal const string User = @"Software\iron-jay\Macro Polo";

        /// <summary>
        /// Returns the first value found for <paramref name="valueName"/>, or null when the
        /// setting is configured nowhere.
        /// </summary>
        internal static object Resolve(IRegistryValueSource registry, string valueName)
        {
            return registry.GetValue(RegistryRoot.LocalMachine, Policy, valueName)
                ?? registry.GetValue(RegistryRoot.LocalMachine, Machine, valueName)
                ?? registry.GetValue(RegistryRoot.CurrentUser, User, valueName);
        }
    }
}
