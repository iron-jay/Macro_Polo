using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace Macro_Polo.Core
{
    /// <summary>
    /// Everything both add-ins do as a native COM shared add-in: lifetime, the ribbon toggle, and
    /// obtaining the task pane factory.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This replaces the VSTO <c>ThisAddIn</c> and ribbon base. The move off VSTO removes ClickOnce
    /// from the picture entirely, and with it the requirement to sign deployment manifests - a
    /// COM add-in is registered from the registry and has no manifests to sign.
    /// </para>
    /// <para>
    /// Office calls the members here in a fixed order: <see cref="OnConnection"/>, then
    /// <see cref="CTPFactoryAvailable"/>, then <see cref="GetCustomUI"/>. The controller is built
    /// when the factory arrives, because it cannot show anything without it.
    /// </para>
    /// <para>
    /// The <see cref="ComVisibleAttribute"/> below is load-bearing and must stay. Both add-in
    /// assemblies are ComVisible(false) at assembly level, so without it this base class is
    /// invisible to COM - and the CLR then refuses QueryInterface for IDispatch on the derived
    /// classes, however visible those are themselves ("this type has a ComVisible(false) parent in
    /// its hierarchy"). IDispatch is how Office invokes every ribbon callback, and losing it takes
    /// the host process down rather than merely disabling the ribbon.
    /// </para>
    /// </remarks>
    [ComVisible(true)]
    public abstract class ComAddInBase : IDTExtensibility2, Office.IRibbonExtensibility, Office.ICustomTaskPaneConsumer
    {
        /// <summary>Must match the toggle's id in the ribbon XML.</summary>
        private const string ToggleId = "Polo";

        private Office.IRibbonUI _ribbon;
        private BannerPaneManager _panes;

        /// <summary>The host <c>Application</c> object, available from <see cref="OnConnection"/> onwards.</summary>
        protected object Application { get; private set; }

        /// <summary>The controller, available from <see cref="CTPFactoryAvailable"/> onwards.</summary>
        protected MacroPoloController Controller { get; private set; }

        #region IDTExtensibility2

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            try
            {
                // Reports exceptions that are caught and handled further up, which would otherwise
                // leave no trace. Only attached when logging is on, so it costs nothing normally.
                if (Log.IsEnabled)
                {
                    AppDomain.CurrentDomain.FirstChanceException += (sender, e) =>
                        Log.Info("first-chance " + e.Exception.GetType().Name + ": " + e.Exception.Message);
                }

                Application = application;
                Log.Info("OnConnection (" + connectMode + "), host=" + (application == null ? "null" : application.GetType().FullName));
            }
            catch (Exception ex)
            {
                // Office disables an add-in that throws out of a lifetime callback, so failures are
                // logged and swallowed: a Macro Polo that does nothing beats one Office refuses to load.
                Log.Error("OnConnection failed", ex);
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            try
            {
                DisconnectHostEvents();
            }
            catch (Exception ex)
            {
                Log.Warn("Failed to detach host event handlers", ex);
            }

            if (Controller != null)
            {
                Controller.Dispose();
                Controller = null;
            }

            _panes = null;
            _ribbon = null;
            Application = null;

            Log.Info("Disconnected (" + removeMode + ")");
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
        }

        #endregion

        #region ICustomTaskPaneConsumer

        public void CTPFactoryAvailable(Office.ICTPFactory CTPFactoryInst)
        {
            try
            {
                if (CTPFactoryInst == null)
                {
                    Log.Error("Office supplied no task pane factory; the banner cannot be shown.", null);
                    return;
                }

                Log.Info("CTPFactoryAvailable");

                _panes = new BannerPaneManager(CTPFactoryInst);

                Controller = CreateController(Application, _panes);
                Log.Info("Controller created");

                Controller.Start();

                ConnectHostEvents();
                Log.Info("Host events connected; startup complete");
            }
            catch (Exception ex)
            {
                Log.Error("Macro Polo failed to start", ex);
            }
        }

        #endregion

        #region IRibbonExtensibility

        public string GetCustomUI(string RibbonID)
        {
            Log.Info("GetCustomUI (" + RibbonID + ")");
            string xml = GetResourceText(RibbonXmlResourceName);
            Log.Info("Ribbon XML " + (xml == null ? "NOT FOUND" : xml.Length + " chars"));
            return xml;
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        public void ButtonClick(Office.IRibbonControl control, bool pressed)
        {
            Log.Info("ButtonClick (pressed=" + pressed + ")");

            if (Controller != null)
            {
                Controller.ToggleForActiveDocument();
            }

            // The banner can refuse to appear (no document, or a failure), so the toggle's state
            // comes from the pane rather than from the click.
            InvalidateToggle();
        }

        public bool ButtonPressed(Office.IRibbonControl control)
        {
            return Controller != null && Controller.IsBannerVisibleForActiveDocument();
        }

        /// <summary>
        /// Ribbon images must be handed over as <c>IPictureDisp</c>. VSTO used to accept a
        /// <see cref="Bitmap"/> and convert on our behalf; talking to Office directly, we convert.
        /// </summary>
        public stdole.IPictureDisp ButtonImage(Office.IRibbonControl control)
        {
            try
            {
                return PictureConverter.ToPictureDisp(ButtonBitmap);
            }
            catch (Exception ex)
            {
                Log.Warn("Could not convert the ribbon image", ex);
                return null;
            }
        }

        #endregion

        /// <summary>Asks Office to re-read the toggle's pressed state.</summary>
        public void InvalidateToggle()
        {
            if (_ribbon == null)
            {
                return;
            }

            try
            {
                _ribbon.InvalidateControl(ToggleId);
            }
            catch (Exception ex)
            {
                // Office discards the IRibbonUI when the window it belongs to goes away.
                Log.Warn("Could not invalidate the ribbon toggle", ex);
            }
        }

        /// <summary>Builds the host-specific controller.</summary>
        protected abstract MacroPoloController CreateController(object application, BannerPaneManager panes);

        /// <summary>Subscribes to the host's document events.</summary>
        protected abstract void ConnectHostEvents();

        /// <summary>Unsubscribes from the host's document events.</summary>
        protected abstract void DisconnectHostEvents();

        /// <summary>Manifest resource name of this host's ribbon XML.</summary>
        protected abstract string RibbonXmlResourceName { get; }

        /// <summary>This host's ribbon button image.</summary>
        protected abstract Bitmap ButtonBitmap { get; }

        private static string GetResourceText(string resourceName)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();

            foreach (string candidate in assembly.GetManifestResourceNames())
            {
                if (!string.Equals(resourceName, candidate, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                using (Stream stream = assembly.GetManifestResourceStream(candidate))
                {
                    if (stream == null)
                    {
                        break;
                    }

                    using (var reader = new StreamReader(stream))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }

            Log.Error("Ribbon XML resource not found: " + resourceName, null);
            return null;
        }
    }
}
