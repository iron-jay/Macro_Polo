using System;
using System.Globalization;
using System.Runtime.InteropServices;

namespace Macro_Polo.Core
{
    /// <summary>Stable identity for a COM object.</summary>
    /// <remarks>
    /// Office hands out a fresh runtime callable wrapper on almost every call, so two references
    /// to the same open document are usually different .NET objects. Using them directly as
    /// dictionary keys quietly fails to match, and it pins the document's RCW alive for as long
    /// as the dictionary holds it. Comparing the underlying IUnknown pointer is the supported way
    /// to ask whether two references are the same COM object.
    /// </remarks>
    public static class ComIdentity
    {
        /// <summary>
        /// Returns a key that is equal for any two references to the same COM object, or null when
        /// <paramref name="comObject"/> is null.
        /// </summary>
        /// <remarks>
        /// The pointer is released immediately: the value is used purely as a key, and holding the
        /// reference is exactly the leak this is meant to avoid. The key is therefore only
        /// meaningful while the caller independently keeps the object alive, which the host does
        /// for as long as the document is open.
        /// </remarks>
        public static string KeyFor(object comObject)
        {
            if (comObject == null)
            {
                return null;
            }

            IntPtr unknown = IntPtr.Zero;
            try
            {
                unknown = Marshal.GetIUnknownForObject(comObject);
                return unknown.ToInt64().ToString("x", CultureInfo.InvariantCulture);
            }
            catch (Exception ex)
            {
                Log.Warn("Could not obtain COM identity", ex);
                return null;
            }
            finally
            {
                if (unknown != IntPtr.Zero)
                {
                    Marshal.Release(unknown);
                }
            }
        }
    }
}
