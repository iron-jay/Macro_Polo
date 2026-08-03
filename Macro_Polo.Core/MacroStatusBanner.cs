using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Macro_Polo.Core
{
    /// <summary>
    /// The strip shown under the ribbon. Renders a <see cref="MacroStatusView"/>.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The control is COM-visible with a fixed CLSID and ProgID because Office's task pane factory
    /// instantiates the hosted control by ProgID. The GUID must not be changed once shipped: the
    /// installer registers it, and an installed machine would be left with a registration pointing
    /// at a class that no longer exists.
    /// </para>
    /// <para>
    /// Sizing is worked out here rather than left to WinForms. <c>AutoScaleMode.Dpi</c> does
    /// nothing in this control: it scales relative to <c>AutoScaleDimensions</c>, which only the
    /// designer sets, and this control is built in code and created by COM. So the scale factor is
    /// derived from the monitor the control is actually on, and the height the host should give the
    /// pane is measured from the laid-out content and published through
    /// <see cref="PreferredHeightChanged"/>.
    /// </para>
    /// </remarks>
    [ComVisible(true)]
    [Guid(Clsid)]
    [ProgId(ProgId)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class MacroStatusBanner : UserControl
    {
        /// <summary>Class id registered for this control. Referenced by the installer.</summary>
        public const string Clsid = "D2E7B4A6-3C81-4F52-9E07-5B6A1C8D4F20";

        /// <summary>ProgID Office uses to create the control. Referenced by the installer.</summary>
        public const string ProgId = "Macro_Polo.StatusBanner";

        /// <summary>Layout measurements, in logical pixels at 96 DPI.</summary>
        private const int LogicalPaddingX = 10;
        private const int LogicalPaddingY = 6;
        private const int LogicalGlyphGap = 10;
        private const int LogicalMinimumHeight = 44;

        /// <summary>How much larger the glyph is than the body text, in points.</summary>
        private const float GlyphPointsLarger = 6f;

        private readonly TableLayoutPanel _layout;
        private readonly TableLayoutPanel _text;
        private readonly Label _glyph;
        private readonly Label _title;
        private readonly Label _detail;
        private readonly ToolTip _toolTip;

        private readonly string _baseFontFamily;
        private readonly float _baseFontPoints;

        /// <summary>
        /// The DPI the system font was sized for. The scale factor is measured against this rather
        /// than against 96, because the system font already reflects the system-wide scaling - so
        /// scaling it again by DeviceDpi/96 would double-apply it on the primary monitor.
        /// </summary>
        private int _baselineDpi;

        private int _appliedDpi;
        private int _publishedHeight = -1;
        private bool _measuring;

        public MacroStatusBanner()
        {
            Font baseFont = SystemFonts.MessageBoxFont;
            _baseFontFamily = baseFont.FontFamily.Name;
            _baseFontPoints = baseFont.SizeInPoints;

            _toolTip = new ToolTip { AutoPopDelay = 30000, InitialDelay = 400, ReshowDelay = 100 };

            _glyph = new Label
            {
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleCenter,
                Margin = new Padding(0)
            };

            _title = new Label { AutoSize = true, Margin = new Padding(0) };
            _detail = new Label { AutoSize = true, Margin = new Padding(0) };

            _text = new TableLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 1,
                RowCount = 2,
                Dock = DockStyle.Fill,
                Margin = new Padding(0)
            };
            _text.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _text.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _text.Controls.Add(_title, 0, 0);
            _text.Controls.Add(_detail, 0, 1);

            _layout = new TableLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2,
                RowCount = 1,
                Dock = DockStyle.Fill,
                Margin = new Padding(0)
            };
            _layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            _layout.Controls.Add(_glyph, 0, 0);
            _layout.Controls.Add(_text, 1, 0);

            Controls.Add(_layout);

            // Screen readers announce a control with the Alert role when it appears, which is the
            // whole point of a banner that shows up on its own when a document is opened.
            AccessibleRole = AccessibleRole.Alert;

            _baselineDpi = DeviceDpi;
            ApplyScale();
        }

        /// <summary>
        /// Raised when the height the banner needs has changed - because the text changed, the pane
        /// was made narrower and the text rewrapped, or the window moved to a monitor with a
        /// different scale factor. The host resizes the task pane in response.
        /// </summary>
        public event EventHandler PreferredHeightChanged;

        /// <summary>The status currently on display, or null before the first update.</summary>
        public MacroStatusView View { get; private set; }

        /// <summary>The height the hosting task pane should be given, at the current width and DPI.</summary>
        public int PreferredPaneHeight { get; private set; }

        /// <summary>Replaces the banner's contents.</summary>
        public void Update(MacroStatusView view)
        {
            if (view == null) throw new ArgumentNullException("view");

            View = view;

            SuspendLayout();
            try
            {
                BackColor = view.BackColor;

                _glyph.Text = view.Glyph;
                _glyph.ForeColor = view.ForeColor;
                _glyph.AccessibleRole = AccessibleRole.Graphic;
                _glyph.AccessibleName = view.Severity.ToString();

                _title.Text = view.Title;
                _title.ForeColor = view.ForeColor;

                _detail.Text = view.Detail;
                _detail.ForeColor = view.ForeColor;

                AccessibleName = view.Title;
                AccessibleDescription = view.Detail;

                _toolTip.SetToolTip(this, view.ToolTip);
                _toolTip.SetToolTip(_glyph, view.ToolTip);
                _toolTip.SetToolTip(_title, view.ToolTip);
                _toolTip.SetToolTip(_detail, view.ToolTip);
            }
            finally
            {
                ResumeLayout(true);
            }

            Measure();
        }

        /// <summary>
        /// Measures the banner for <paramref name="availableWidth"/> and returns the height the
        /// pane needs. Kept for the host's first call, before the control has been given a size.
        /// </summary>
        public int GetPreferredPaneHeight(int availableWidth)
        {
            Measure(availableWidth);
            return PreferredPaneHeight;
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            Measure();
        }

        /// <summary>
        /// Fires when the pane is dragged to a monitor with a different scale factor. WinForms does
        /// not rescale a control's fonts by itself here, so the banner does it.
        /// </summary>
        protected override void OnDpiChangedAfterParent(EventArgs e)
        {
            base.OnDpiChangedAfterParent(e);
            ApplyScale();
            Measure();
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);

            // DeviceDpi is only meaningful once there is a handle. If the control was constructed
            // before Office sited it, the baseline captured in the constructor was a guess.
            if (_baselineDpi <= 0)
            {
                _baselineDpi = DeviceDpi;
            }

            ApplyScale();
            Measure();
        }

        /// <summary>
        /// Converts a logical (96 DPI) pixel measurement to device pixels. Padding, margins and
        /// minimum heights are all in pixels, so they scale with the monitor absolutely.
        /// </summary>
        private float PixelScale
        {
            get { return (DeviceDpi > 0 ? DeviceDpi : 96) / 96f; }
        }

        /// <summary>
        /// Correction applied to font sizes, which is a different thing entirely from
        /// <see cref="PixelScale"/> and must not be confused with it.
        /// </summary>
        /// <remarks>
        /// Font sizes are in points, and points are already resolution independent: the system font
        /// is 9pt whether the display is at 100% or 150%, and rendering it on a 150% monitor
        /// produces bigger glyphs on its own. Multiplying the point size by DeviceDpi/96 as well
        /// would apply the scaling twice and give enormous text. So this stays at 1 on the monitor
        /// the font was measured for, and only corrects the difference when the window is dragged
        /// to a monitor with a different scale factor.
        /// </remarks>
        private float FontScale
        {
            get
            {
                int dpi = DeviceDpi > 0 ? DeviceDpi : 96;
                int baseline = _baselineDpi > 0 ? _baselineDpi : 96;
                return (float)dpi / baseline;
            }
        }

        /// <summary>Rebuilds fonts and spacing for the monitor the control is currently on.</summary>
        private void ApplyScale()
        {
            if (_appliedDpi == DeviceDpi)
            {
                return;
            }

            _appliedDpi = DeviceDpi;

            float pixels = PixelScale;
            float points = FontScale;

            SuspendLayout();
            try
            {
                Padding = new Padding(
                    Round(LogicalPaddingX * pixels),
                    Round(LogicalPaddingY * pixels),
                    Round(LogicalPaddingX * pixels),
                    Round(LogicalPaddingY * pixels));

                _glyph.Margin = new Padding(0, 0, Round(LogicalGlyphGap * pixels), 0);
                _detail.Margin = new Padding(0, Round(2 * pixels), 0, 0);

                ReplaceFont(_glyph, new Font("Segoe UI Symbol", (_baseFontPoints + GlyphPointsLarger) * points, FontStyle.Regular));
                ReplaceFont(_title, new Font(_baseFontFamily, _baseFontPoints * points, FontStyle.Bold));
                ReplaceFont(_detail, new Font(_baseFontFamily, _baseFontPoints * points, FontStyle.Regular));
            }
            finally
            {
                ResumeLayout(true);
            }
        }

        private static void ReplaceFont(Control control, Font font)
        {
            Font old = control.Font;
            control.Font = font;

            // Never dispose the ambient font the control started with.
            if (old != null && old != SystemFonts.MessageBoxFont && old != Control.DefaultFont)
            {
                old.Dispose();
            }
        }

        private void Measure()
        {
            Measure(ClientSize.Width);
        }

        /// <summary>
        /// Sets the wrapping width, measures the laid-out content, and publishes the height if it
        /// changed.
        /// </summary>
        private void Measure(int availableWidth)
        {
            // Setting the pane height causes a resize, which lands back here. Without this guard
            // that becomes a feedback loop between the banner and the task pane.
            if (_measuring)
            {
                return;
            }

            _measuring = true;
            try
            {
                int textWidth = TextWidthFor(availableWidth);
                if (textWidth <= 0)
                {
                    return;
                }

                // Labels only wrap when given an explicit maximum width.
                _title.MaximumSize = new Size(textWidth, 0);
                _detail.MaximumSize = new Size(textWidth, 0);

                int content = _layout.GetPreferredSize(new Size(textWidth, 0)).Height + Padding.Vertical;
                int measured = Math.Max(content, Round(LogicalMinimumHeight * PixelScale));

                PreferredPaneHeight = measured;

                if (measured != _publishedHeight)
                {
                    _publishedHeight = measured;

                    EventHandler handler = PreferredHeightChanged;
                    if (handler != null)
                    {
                        handler(this, EventArgs.Empty);
                    }
                }
            }
            finally
            {
                _measuring = false;
            }
        }

        /// <summary>Width available to the text once padding and the glyph are taken out.</summary>
        private int TextWidthFor(int availableWidth)
        {
            int glyphWidth = _glyph.PreferredSize.Width + _glyph.Margin.Horizontal;
            return availableWidth - Padding.Horizontal - glyphWidth;
        }

        private static int Round(float value)
        {
            return (int)Math.Round(value, MidpointRounding.AwayFromZero);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _toolTip.Dispose();
            }

            base.Dispose(disposing);
        }
    }
}
