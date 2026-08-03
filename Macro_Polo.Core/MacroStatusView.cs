using System.Drawing;

namespace Macro_Polo.Core
{
    /// <summary>How alarming a status is. Drives colour and glyph.</summary>
    public enum Severity
    {
        /// <summary>Nothing to report.</summary>
        Neutral,

        /// <summary>Macros are present and running, and they are signed.</summary>
        Safe,

        /// <summary>Macros are present but cannot run right now.</summary>
        Information,

        /// <summary>Macros are present and the user can choose to run them.</summary>
        Caution,

        /// <summary>Unsigned macros are running, or will run, with no prompt.</summary>
        Danger
    }

    /// <summary>Everything the banner needs in order to render itself.</summary>
    public sealed class MacroStatusView
    {
        public MacroStatusView(Severity severity, string title, string detail, string toolTip)
        {
            Severity = severity;
            Title = title;
            Detail = detail;
            ToolTip = toolTip;
        }

        public Severity Severity { get; private set; }

        public string Title { get; private set; }

        public string Detail { get; private set; }

        /// <summary>The long form, including the caveats that will not fit in the banner.</summary>
        public string ToolTip { get; private set; }

        public Color BackColor
        {
            get { return Palette.Background(Severity); }
        }

        public Color ForeColor
        {
            get { return Palette.Foreground(Severity); }
        }

        /// <summary>
        /// A glyph carrying the same meaning as the colour, so the status is not conveyed by
        /// colour alone.
        /// </summary>
        public string Glyph
        {
            get
            {
                switch (Severity)
                {
                    case Severity.Safe: return "✔";        // heavy check mark
                    case Severity.Caution: return "⚠";     // warning sign
                    case Severity.Danger: return "✖";      // heavy multiplication x
                    case Severity.Information: return "⛔"; // no entry
                    default: return "●";                   // black circle
                }
            }
        }
    }

    /// <summary>
    /// The banner colours. Every pair clears WCAG AA contrast for the bold 12pt text used in the
    /// banner.
    /// </summary>
    internal static class Palette
    {
        internal static Color Background(Severity severity)
        {
            switch (severity)
            {
                case Severity.Safe: return ColorTranslator.FromHtml("#225D2E");
                case Severity.Information: return ColorTranslator.FromHtml("#205493");
                case Severity.Caution: return ColorTranslator.FromHtml("#F9C642");
                case Severity.Danger: return ColorTranslator.FromHtml("#981B1E");
                default: return ColorTranslator.FromHtml("#323A45");
            }
        }

        internal static Color Foreground(Severity severity)
        {
            return severity == Severity.Caution
                ? ColorTranslator.FromHtml("#212121")
                : ColorTranslator.FromHtml("#FFFFFF");
        }
    }
}
