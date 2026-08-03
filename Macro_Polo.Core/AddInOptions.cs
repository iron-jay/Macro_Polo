using System;

namespace Macro_Polo.Core
{
    /// <summary>When the banner should appear without being asked for.</summary>
    public enum AutoShowMode
    {
        /// <summary>Only show the banner when the ribbon button is pressed. The default.</summary>
        Never = 0,

        /// <summary>Show the banner automatically when the document contains macros.</summary>
        WhenMacrosPresent = 1,

        /// <summary>Show the banner for every document, including those with no macros.</summary>
        Always = 2
    }

    /// <summary>
    /// Add-in behaviour, configurable so that an administrator can set it for a fleet rather than
    /// relying on every user to press a button.
    /// </summary>
    /// <remarks>See <see cref="OptionScopes"/> for where these are read from, and in what order.</remarks>
    public sealed class AddInOptions
    {
        public AddInOptions()
        {
            AutoShow = AutoShowMode.Never;
        }

        public AutoShowMode AutoShow { get; set; }

        public static AddInOptions Read(IRegistryValueSource registry)
        {
            if (registry == null) throw new ArgumentNullException("registry");

            var options = new AddInOptions();

            object raw = OptionScopes.Resolve(registry, "AutoShow");

            if (raw != null)
            {
                try
                {
                    int value = Convert.ToInt32(raw, System.Globalization.CultureInfo.InvariantCulture);
                    if (Enum.IsDefined(typeof(AutoShowMode), value))
                    {
                        options.AutoShow = (AutoShowMode)value;
                    }
                }
                catch (Exception ex)
                {
                    Log.Warn("Ignoring unreadable AutoShow value", ex);
                }
            }

            return options;
        }
    }
}
