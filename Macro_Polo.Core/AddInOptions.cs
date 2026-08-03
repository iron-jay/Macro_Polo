using System;

namespace Macro_Polo.Core
{
    /// <summary>When the banner should appear without being asked for.</summary>
    public enum AutoShowMode
    {
        /// <summary>Only ever show the banner when the ribbon button is pressed.</summary>
        Never = 0,

        /// <summary>Show the banner automatically when the document contains macros. The default.</summary>
        WhenMacrosPresent = 1,

        /// <summary>Show the banner for every document, including those with no macros.</summary>
        Always = 2
    }

    /// <summary>
    /// Add-in behaviour, configurable so that an administrator can set it for a fleet rather than
    /// relying on every user to press a button.
    /// </summary>
    /// <remarks>
    /// Read from, in order of precedence: <c>HKLM\Software\Policies\Macro_Polo</c>,
    /// <c>HKLM\Software\Macro_Polo</c>, <c>HKCU\Software\Macro_Polo</c>.
    /// </remarks>
    public sealed class AddInOptions
    {
        private const string PolicyPath = @"Software\Policies\Macro_Polo";
        private const string MachinePath = @"Software\Macro_Polo";
        private const string UserPath = @"Software\Macro_Polo";

        public AddInOptions()
        {
            AutoShow = AutoShowMode.WhenMacrosPresent;
        }

        public AutoShowMode AutoShow { get; set; }

        public static AddInOptions Read(IRegistryValueSource registry)
        {
            if (registry == null) throw new ArgumentNullException("registry");

            var options = new AddInOptions();

            object raw = registry.GetValue(RegistryRoot.LocalMachine, PolicyPath, "AutoShow")
                ?? registry.GetValue(RegistryRoot.LocalMachine, MachinePath, "AutoShow")
                ?? registry.GetValue(RegistryRoot.CurrentUser, UserPath, "AutoShow");

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
