using System;
using System.Collections.Generic;

namespace Macro_Polo.Core
{
    /// <summary>
    /// All of the add-in's behaviour that does not depend on whether the host is Word or Excel.
    /// A host supplies three things: how to find the active document, how to find its window, and
    /// how to read the macro facts off it.
    /// </summary>
    /// <remarks>
    /// This type is compiled into both add-in projects as a linked source file rather than living
    /// in Macro_Polo.Core, because it reaches the Office task pane types through
    /// <see cref="BannerPaneManager"/> and Macro_Polo.Core is kept free of Office references so it
    /// can be built and tested anywhere.
    /// </remarks>
    public abstract class MacroPoloController : IDisposable
    {
        private readonly BannerPaneManager _panes;
        private readonly IRegistryValueSource _registry;
        private readonly string _applicationRegistryName;
        private readonly object _application;

        /// <summary>
        /// Documents the banner has already been offered for. Without this, re-activating a
        /// document would keep reopening a banner the user has deliberately closed.
        /// </summary>
        private readonly HashSet<string> _autoShown = new HashSet<string>(StringComparer.Ordinal);

        private AddInOptions _options;
        private string _officeVersion;
        private bool _disposed;

        /// <param name="panes">Banner task panes, created from Office's task pane factory.</param>
        /// <param name="application">The host <c>Application</c> object.</param>
        /// <param name="applicationRegistryName">"Word" or "Excel", as Office spells it in the registry.</param>
        protected MacroPoloController(BannerPaneManager panes, object application, string applicationRegistryName)
        {
            if (panes == null) throw new ArgumentNullException("panes");

            _application = application;
            _applicationRegistryName = applicationRegistryName;
            _registry = new WindowsRegistryValueSource();
            _panes = panes;
        }

        public void Start()
        {
            _options = AddInOptions.Read(_registry);
            _officeVersion = OfficeVersion.FromHost(_application);

            Log.Info("Macro Polo started for " + _applicationRegistryName + " " + _officeVersion
                + " (auto-show: " + _options.AutoShow + ")");
        }

        /// <summary>
        /// Called when a document is opened or brought to the front. Shows the banner by itself,
        /// which is the point: a user who does not know a file has macros also does not know to
        /// press a button asking.
        /// </summary>
        public void HandleDocumentPresented(object document)
        {
            if (_disposed || document == null)
            {
                return;
            }

            try
            {
                string key = ComIdentity.KeyFor(document);
                if (key == null)
                {
                    return;
                }

                Log.Info("HandleDocumentPresented key=" + key);

                MacroStatus status = Evaluate(document);
                Log.Info("Evaluated: " + status.State);

                MacroStatusView view = MacroStatusPresenter.Describe(status);
                Log.Info("Described: " + view.Severity);

                if (_panes.IsVisible(key))
                {
                    _panes.Update(key, view);
                    return;
                }

                if (_autoShown.Contains(key) || !ShouldAutoShow(status))
                {
                    Log.Info("Not auto-showing (already offered, or not wanted for this state)");
                    return;
                }

                _autoShown.Add(key);
                _panes.Show(key, GetWindow(document), view);
            }
            catch (Exception ex)
            {
                // Never let a status check propagate into Office: an add-in that throws gets
                // disabled, and the user is told nothing useful about why.
                Log.Error("Failed to present macro status", ex);
            }
        }

        /// <summary>Ribbon button. Always acts, including for documents with no macros.</summary>
        public void ToggleForActiveDocument()
        {
            if (_disposed)
            {
                return;
            }

            try
            {
                object document = GetActiveDocument();
                if (document == null)
                {
                    ShowNoDocumentMessage();
                    return;
                }

                string key = ComIdentity.KeyFor(document);
                if (key == null)
                {
                    return;
                }

                _autoShown.Add(key);
                _panes.Toggle(key, GetWindow(document), MacroStatusPresenter.Describe(Evaluate(document)));
            }
            catch (Exception ex)
            {
                Log.Error("Failed to toggle macro status banner", ex);
                ShowCheckFailedMessage(ex);
            }
        }

        /// <summary>True when the active document's banner is currently on screen.</summary>
        public bool IsBannerVisibleForActiveDocument()
        {
            try
            {
                object document = GetActiveDocument();
                return document != null && _panes.IsVisible(ComIdentity.KeyFor(document));
            }
            catch (Exception ex)
            {
                Log.Warn("Could not determine banner visibility", ex);
                return false;
            }
        }

        /// <summary>
        /// Called as a document closes. Releases its pane; a pane that is only hidden stays in the
        /// collection for the rest of the session.
        /// </summary>
        public void HandleDocumentClosed(object document)
        {
            if (document == null)
            {
                return;
            }

            try
            {
                string key = ComIdentity.KeyFor(document);
                if (key == null)
                {
                    return;
                }

                // Word and Excel both raise this before the close can be cancelled, so forget that
                // the banner was already offered: if the close is called off, the next activation
                // should put the banner back.
                _autoShown.Remove(key);
                _panes.Close(key);
            }
            catch (Exception ex)
            {
                Log.Warn("Failed to release banner for a closing document", ex);
            }
        }

        private bool ShouldAutoShow(MacroStatus status)
        {
            switch (_options.AutoShow)
            {
                case AutoShowMode.Always:
                    return true;
                case AutoShowMode.WhenMacrosPresent:
                    return status.Document.HasMacros;
                default:
                    return false;
            }
        }

        /// <summary>
        /// Reads the Trust Center configuration afresh on every check. It is a handful of registry
        /// reads, and caching it means reporting a stale answer after the user or a policy refresh
        /// changes the macro setting mid-session.
        /// </summary>
        private MacroStatus Evaluate(object document)
        {
            var reader = new OfficeSecurityReader(_registry, _officeVersion ?? OfficeVersion.Fallback, _applicationRegistryName);
            DocumentMacroInfo info = Describe(document);
            info.HasMarkOfTheWeb = MarkOfTheWeb.IsPresent(info.FullPath);

            return MacroStatusEvaluator.Evaluate(info, reader.Read());
        }

        private object GetWindow(object document)
        {
            try
            {
                return GetWindowFor(document);
            }
            catch (Exception ex)
            {
                Log.Warn("Could not resolve the document window; the pane will use the active one", ex);
                return null;
            }
        }

        /// <summary>The document currently in front, or null when the host has none open.</summary>
        protected abstract object GetActiveDocument();

        /// <summary>The window <paramref name="document"/> is displayed in.</summary>
        protected abstract object GetWindowFor(object document);

        /// <summary>Reads the macro facts off a host document.</summary>
        protected abstract DocumentMacroInfo Describe(object document);

        /// <summary>Told the user there is nothing to check.</summary>
        protected abstract void ShowNoDocumentMessage();

        /// <summary>Told the user the check itself failed.</summary>
        protected abstract void ShowCheckFailedMessage(Exception exception);

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            _panes.Dispose();
            _autoShown.Clear();
        }
    }
}
