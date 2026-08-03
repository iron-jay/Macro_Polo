using System;
using System.Collections.Generic;
using System.Linq;
using Macro_Polo.Core;

namespace Macro_Polo.Core.Tests
{
    /// <summary>
    /// An in-memory stand-in for the registry, so the hive precedence rules can be exercised
    /// without touching the machine.
    /// </summary>
    internal sealed class FakeRegistry : IRegistryValueSource
    {
        private readonly Dictionary<string, object> _values =
            new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

        internal FakeRegistry Set(RegistryRoot root, string subKeyPath, string valueName, object value)
        {
            _values[Key(root, subKeyPath, valueName)] = value;
            return this;
        }

        public object GetValue(RegistryRoot root, string subKeyPath, string valueName)
        {
            object value;
            return _values.TryGetValue(Key(root, subKeyPath, valueName), out value) ? value : null;
        }

        public IEnumerable<string> GetSubKeyNames(RegistryRoot root, string subKeyPath)
        {
            string prefix = root + "|" + subKeyPath + "\\";

            return _values.Keys
                .Where(k => k.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                .Select(k => k.Substring(prefix.Length))
                // What remains looks like "Location0|Path": strip the value name, then take the
                // first path segment so only immediate children are returned.
                .Select(rest => rest.Split('|')[0].Split('\\')[0])
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static string Key(RegistryRoot root, string subKeyPath, string valueName)
        {
            return root + "|" + subKeyPath + "|" + valueName;
        }
    }
}
