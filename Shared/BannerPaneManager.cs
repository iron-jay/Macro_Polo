using System;
using System.Collections.Generic;
using Office = Microsoft.Office.Core;

namespace Macro_Polo.Core
{
    /// <summary>
    /// Owns one banner task pane per document window.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Panes are created through Office's own <c>ICTPFactory</c> rather than VSTO's task pane
    /// collection. The factory instantiates the hosted control by ProgID, which is why
    /// <see cref="MacroStatusBanner"/> is registered as a COM control.
    /// </para>
    /// <para>
    /// Each pane is created against the window it belongs to. Without that, Office attaches the
    /// pane to whichever window happens to be active at creation time, and the banner then shows
    /// up over the wrong document.
    /// </para>
    /// </remarks>
    public sealed class BannerPaneManager : IDisposable
    {
        private readonly Office.ICTPFactory _factory;
        private readonly Dictionary<string, Entry> _entries = new Dictionary<string, Entry>(StringComparer.Ordinal);
        private bool _disposed;

        public BannerPaneManager(Office.ICTPFactory factory)
        {
            if (factory == null) throw new ArgumentNullException("factory");
            _factory = factory;
        }

        public bool IsVisible(string key)
        {
            Entry entry;
            if (key == null || !_entries.TryGetValue(key, out entry))
            {
                return false;
            }

            try
            {
                return entry.Pane.Visible;
            }
            catch (Exception ex)
            {
                // The underlying pane can already be gone if Office tore down the window.
                Log.Warn("Could not read pane visibility", ex);
                Forget(key);
                return false;
            }
        }

        /// <summary>Creates or updates the pane for <paramref name="key"/> and makes it visible.</summary>
        public void Show(string key, object window, MacroStatusView view)
        {
            if (key == null || view == null || _disposed)
            {
                return;
            }

            try
            {
                Entry entry;
                if (_entries.TryGetValue(key, out entry))
                {
                    entry.Banner.Update(view);
                    entry.Pane.Visible = true;
                    ResizeToContent(entry);
                    return;
                }

                // Each COM boundary below is logged separately. All of them are calls whose failure
                // mode is a host process crash rather than a managed exception, so when something
                // does go wrong the last line in the log is what identifies it.
                Log.Info("CreateCTP: progId=" + MacroStatusBanner.ProgId + " window=" + (window == null ? "<none>" : window.GetType().Name));

                // CreateCTP treats a missing parent as "the active window"; it will not accept null.
                //
                // The result is held as _CustomTaskPane, not CustomTaskPane. The latter additionally
                // carries the event interface, and binding a COM event sink to it through embedded
                // interop types is the least reliable construct in this file. Nothing here needs the
                // events: the ribbon toggle reads Visible on demand.
                Office._CustomTaskPane pane = _factory.CreateCTP(
                    MacroStatusBanner.ProgId,
                    Strings.Get("PaneTitle"),
                    window ?? Type.Missing);

                Log.Info("CreateCTP returned " + (pane == null ? "null" : "a pane"));
                if (pane == null)
                {
                    return;
                }

                // ContentControl hands back the control the factory created. It is our own .NET
                // object behind a COM wrapper, so the cast unwraps to the original instance.
                object content = pane.ContentControl;
                Log.Info("ContentControl is " + (content == null ? "null" : content.GetType().FullName));

                var banner = content as MacroStatusBanner;
                if (banner == null)
                {
                    Log.Error("The task pane did not host a Macro Polo banner. Is the control registered?", null);
                    pane.Delete();
                    return;
                }

                banner.Update(view);
                Log.Info("Banner content set; docking");

                pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionTop;

                entry = new Entry(pane, banner, key);
                _entries.Add(key, entry);

                // From here on the pane height follows the banner. That covers the text rewrapping
                // when the pane is made narrower, and the scale factor changing when the window is
                // dragged to another monitor - neither of which a one-off measurement at creation
                // time can account for.
                entry.Attach(ResizeToContent);

                Log.Info("Docked; making visible");
                pane.Visible = true;

                ResizeToContent(entry);
                Log.Info("Banner shown");
            }
            catch (Exception ex)
            {
                Log.Error("Could not show the macro status banner", ex);
                Forget(key);
            }
        }

        /// <summary>Refreshes an existing pane in place. Does nothing if there is no pane.</summary>
        public void Update(string key, MacroStatusView view)
        {
            Entry entry;
            if (key == null || view == null || !_entries.TryGetValue(key, out entry))
            {
                return;
            }

            try
            {
                entry.Banner.Update(view);
                ResizeToContent(entry);
            }
            catch (Exception ex)
            {
                Log.Warn("Could not update the macro status banner", ex);
            }
        }

        /// <summary>Hides the pane if it is showing, shows it otherwise.</summary>
        public void Toggle(string key, object window, MacroStatusView view)
        {
            if (IsVisible(key))
            {
                Hide(key);
            }
            else
            {
                Show(key, window, view);
            }
        }

        public void Hide(string key)
        {
            Entry entry;
            if (key == null || !_entries.TryGetValue(key, out entry))
            {
                return;
            }

            try
            {
                entry.Pane.Visible = false;
            }
            catch (Exception ex)
            {
                Log.Warn("Could not hide the macro status banner", ex);
                Forget(key);
            }
        }

        /// <summary>
        /// Deletes the pane for a document that is closing. The host must call this: a pane that is
        /// merely hidden stays alive, and its entry would otherwise accumulate for the lifetime of
        /// the Office session.
        /// </summary>
        public void Close(string key)
        {
            Entry entry;
            if (key == null || !_entries.TryGetValue(key, out entry))
            {
                return;
            }

            _entries.Remove(key);
            entry.Dispose();
        }

        /// <summary>
        /// Sizes the pane to the text it is actually showing, at the DPI of the monitor the banner
        /// is on.
        /// </summary>
        private static void ResizeToContent(Entry entry)
        {
            try
            {
                int width = entry.Banner.Width > 0 ? entry.Banner.Width : entry.Pane.Width;
                int content = entry.Banner.GetPreferredPaneHeight(width);

                if (content <= 0)
                {
                    return;
                }

                // A task pane's Height is the height of the whole pane, caption bar included, but
                // the banner only gets what is left after Office has taken its chrome. Setting the
                // pane to the height of the content alone leaves the banner clipped to a sliver.
                //
                // The chrome is measured rather than assumed: its height depends on the Office
                // version, the theme and the monitor's scale factor, so there is no constant worth
                // hard-coding. Until the control has been sited there is nothing to measure, and
                // the first pass simply gets it wrong by the height of the caption - the resize
                // that follows then corrects it, because by then the banner has a real height.
                int chrome = entry.Banner.Height > 0 ? entry.Pane.Height - entry.Banner.Height : 0;
                if (chrome < 0)
                {
                    chrome = 0;
                }

                int wanted = content + chrome;

                // Only assign when it actually differs. Writing the same height back would still
                // provoke a layout pass in Office, and the banner would measure again in response.
                if (wanted != entry.Pane.Height)
                {
                    Log.Info("Pane height " + entry.Pane.Height + " -> " + wanted
                        + " (content " + content + " + chrome " + chrome + ")");

                    entry.Pane.Height = wanted;
                }
            }
            catch (Exception ex)
            {
                Log.Warn("Could not resize the macro status banner", ex);
            }
        }

        /// <summary>Drops an entry whose pane is no longer usable, without touching COM again.</summary>
        private void Forget(string key)
        {
            Entry entry;
            if (key != null && _entries.TryGetValue(key, out entry))
            {
                _entries.Remove(key);
                entry.Abandon();
            }
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;

            foreach (Entry entry in _entries.Values)
            {
                entry.Dispose();
            }

            _entries.Clear();
        }

        /// <summary>
        /// One document's pane. This exists so that the visibility handler closes over a per-pane
        /// object rather than over a field on the add-in, which is what made the original
        /// implementation act on the most recently created pane instead of its own.
        /// </summary>
        private sealed class Entry : IDisposable
        {
            private Action<Entry> _resize;
            private EventHandler _preferredHeightChanged;
            private EventHandler _sizeChanged;
            private bool _resizing;
            private bool _disposed;

            internal Entry(Office._CustomTaskPane pane, MacroStatusBanner banner, string key)
            {
                Pane = pane;
                Banner = banner;
                Key = key;
            }

            /// <summary>
            /// Subscribes to the banner so the pane follows it. The handlers close over this entry
            /// rather than over a field on the manager, so several open documents each resize their
            /// own pane.
            /// </summary>
            /// <remarks>
            /// SizeChanged matters as much as PreferredHeightChanged, and for a reason that is easy
            /// to miss: when the pane is too short, the banner is clipped but the height its content
            /// *wants* has not changed, so PreferredHeightChanged never fires and the correction
            /// pass never happens. Watching the control's actual size catches the case where the
            /// host gave it something other than what it asked for - which is exactly what the
            /// first, chrome-less measurement does.
            /// </remarks>
            internal void Attach(Action<Entry> resize)
            {
                _resize = resize;
                _preferredHeightChanged = (sender, e) => Reresize();
                _sizeChanged = (sender, e) => Reresize();

                Banner.PreferredHeightChanged += _preferredHeightChanged;
                Banner.SizeChanged += _sizeChanged;
            }

            /// <summary>
            /// Runs a resize pass, refusing to re-enter. Setting the pane height resizes the banner,
            /// which lands back here; the pass converges because the manager only assigns a height
            /// that differs from the current one.
            /// </summary>
            private void Reresize()
            {
                if (_resizing || _disposed)
                {
                    return;
                }

                _resizing = true;
                try
                {
                    _resize(this);
                }
                finally
                {
                    _resizing = false;
                }
            }

            internal Office._CustomTaskPane Pane { get; private set; }

            internal MacroStatusBanner Banner { get; private set; }

            internal string Key { get; private set; }

            public void Dispose()
            {
                if (_disposed)
                {
                    return;
                }

                _disposed = true;

                if (_preferredHeightChanged != null)
                {
                    Banner.PreferredHeightChanged -= _preferredHeightChanged;
                    _preferredHeightChanged = null;
                }

                if (_sizeChanged != null)
                {
                    Banner.SizeChanged -= _sizeChanged;
                    _sizeChanged = null;
                }

                try
                {
                    Pane.Delete();
                }
                catch (Exception ex)
                {
                    Log.Warn("Could not delete task pane for " + Key, ex);
                }
            }

            /// <summary>Used when the pane itself is already gone and touching it would throw again.</summary>
            internal void Abandon()
            {
                _disposed = true;
            }
        }
    }
}
