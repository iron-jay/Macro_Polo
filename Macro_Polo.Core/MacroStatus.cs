namespace Macro_Polo.Core
{
    /// <summary>
    /// The result of evaluating one document: the state plus the inputs that produced it, so the
    /// presentation layer can word its message precisely.
    /// </summary>
    public sealed class MacroStatus
    {
        public MacroStatus(MacroState state, DocumentMacroInfo document, MacroSecuritySettings settings)
        {
            State = state;
            Document = document;
            Settings = settings;
        }

        public MacroState State { get; private set; }

        public DocumentMacroInfo Document { get; private set; }

        public MacroSecuritySettings Settings { get; private set; }

        /// <summary>True when the macros will execute without the user having to do anything.</summary>
        public bool RunsWithoutPrompting
        {
            get
            {
                return State == MacroState.RunsSilently
                    || State == MacroState.RunsSilentlyTrustedLocation;
            }
        }

        /// <summary>True when nothing the user does inside Office will let these macros run.</summary>
        public bool IsBlocked
        {
            get
            {
                return State == MacroState.BlockedByPolicy
                    || State == MacroState.BlockedUnsigned
                    || State == MacroState.BlockedFromInternet;
            }
        }
    }
}
