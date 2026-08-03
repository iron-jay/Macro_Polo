using System.Collections.Generic;

namespace Macro_Polo.Core
{
    /// <summary>Which registry root a lookup should target.</summary>
    public enum RegistryRoot
    {
        CurrentUser,
        LocalMachine
    }

    /// <summary>
    /// The narrow slice of the registry that <see cref="OfficeSecurityReader"/> needs. Exists so
    /// the hive precedence rules can be tested without writing to a real machine's registry.
    /// </summary>
    public interface IRegistryValueSource
    {
        /// <summary>
        /// Returns the raw value, or null when either the key or the value is absent.
        /// </summary>
        object GetValue(RegistryRoot root, string subKeyPath, string valueName);

        /// <summary>
        /// Returns the names of the immediate child keys of <paramref name="subKeyPath"/>, or an
        /// empty sequence when the key is absent.
        /// </summary>
        IEnumerable<string> GetSubKeyNames(RegistryRoot root, string subKeyPath);

        /// <summary>
        /// Returns the value names under <paramref name="subKeyPath"/>, or an empty sequence when
        /// the key is absent. Needed because Office records trusted documents as one value per
        /// document, keyed by path, where the name carries the information and not the data.
        /// </summary>
        IEnumerable<string> GetValueNames(RegistryRoot root, string subKeyPath);
    }
}
