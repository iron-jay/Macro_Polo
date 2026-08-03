using System;
using System.Globalization;
using System.Reflection;
using System.Resources;

namespace Macro_Polo.Core
{
    /// <summary>
    /// Localisable UI text. Backed by <c>Resources\Strings.resx</c>; ship a satellite assembly
    /// built from <c>Strings.&lt;culture&gt;.resx</c> to translate the add-in.
    /// </summary>
    /// <remarks>
    /// Lookups are by name rather than through a generated designer class so that the resource
    /// file can be edited without a Visual Studio round trip.
    /// </remarks>
    public static class Strings
    {
        private static readonly ResourceManager Resources = new ResourceManager(
            "Macro_Polo.Core.Resources.Strings",
            Assembly.GetExecutingAssembly());

        /// <summary>
        /// Returns the string for <paramref name="name"/>, falling back to the name itself if the
        /// resource is missing. Missing text should never be able to break the banner.
        /// </summary>
        public static string Get(string name)
        {
            try
            {
                return Resources.GetString(name, CultureInfo.CurrentUICulture) ?? name;
            }
            catch (Exception ex)
            {
                Log.Warn("Missing resource string: " + name, ex);
                return name;
            }
        }

        public static string Format(string name, params object[] args)
        {
            return string.Format(CultureInfo.CurrentUICulture, Get(name), args);
        }

        /// <summary>Human-readable description of a macro setting.</summary>
        public static string Describe(VbaWarningLevel level)
        {
            switch (level)
            {
                case VbaWarningLevel.EnableAll: return Get("Level_EnableAll");
                case VbaWarningLevel.DisableExceptSigned: return Get("Level_DisableExceptSigned");
                case VbaWarningLevel.DisableAll: return Get("Level_DisableAll");
                default: return Get("Level_DisableWithNotification");
            }
        }
    }
}
