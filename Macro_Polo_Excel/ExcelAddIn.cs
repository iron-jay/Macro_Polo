using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Macro_Polo.Core;
using Macro_Polo_Excel.Properties;
using Excel = Microsoft.Office.Interop.Excel;

namespace Macro_Polo_Excel
{
    /// <summary>
    /// The Excel add-in. Office creates this class by ProgID and drives it through
    /// <see cref="IDTExtensibility2"/>.
    /// </summary>
    /// <remarks>
    /// The class id and ProgID are fixed and must not change once shipped: the installer writes
    /// them into the registry, and an installed machine would otherwise be left pointing at a
    /// class that no longer exists. The ProgID doubles as the key name Office looks for under
    /// <c>Software\Microsoft\Office\Excel\Addins</c>.
    /// </remarks>
    [ComVisible(true)]
    [Guid(Clsid)]
    [ProgId(ProgId)]
    // Office resolves ribbon callbacks by name through IDispatch, which AutoDispatch provides by
    // reflecting over the class at run time. AutoDual would additionally bake a static dual vtable
    // out of every public member - including ones typed in embedded interop types - which buys
    // nothing here and is fragile across rebuilds.
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class ExcelAddIn : ComAddInBase
    {
        /// <summary>Class id registered for this add-in. Referenced by the installer.</summary>
        public const string Clsid = "6A4D3C21-8E97-45B2-A0F6-3D7B1E9C5842";

        /// <summary>ProgID, and the name of this add-in's key under Office's Addins key.</summary>
        public const string ProgId = "Macro_Polo.ExcelAddIn";

        private Excel.Application _application;

        protected override string RibbonXmlResourceName
        {
            get { return "Macro_Polo_Excel.Ribbon1.xml"; }
        }

        protected override Bitmap ButtonBitmap
        {
            get { return Resources.excel; }
        }

        protected override MacroPoloController CreateController(object application, BannerPaneManager panes)
        {
            _application = application as Excel.Application;

            if (_application == null)
            {
                throw new InvalidOperationException("Macro Polo was loaded by a host that is not Excel.");
            }

            return new ExcelController(panes, _application);
        }

        protected override void ConnectHostEvents()
        {
            _application.WorkbookOpen += OnWorkbookPresented;
            _application.WorkbookActivate += OnWorkbookPresented;
            _application.WorkbookBeforeClose += OnWorkbookBeforeClose;
        }

        protected override void DisconnectHostEvents()
        {
            if (_application == null)
            {
                return;
            }

            _application.WorkbookOpen -= OnWorkbookPresented;
            _application.WorkbookActivate -= OnWorkbookPresented;
            _application.WorkbookBeforeClose -= OnWorkbookBeforeClose;
            _application = null;
        }

        private void OnWorkbookPresented(Excel.Workbook workbook)
        {
            Controller.HandleDocumentPresented(workbook);
        }

        private void OnWorkbookBeforeClose(Excel.Workbook workbook, ref bool cancel)
        {
            Controller.HandleDocumentClosed(workbook);
        }

        /// <summary>Excel's answers to the host-specific questions.</summary>
        private sealed class ExcelController : MacroPoloController
        {
            private readonly Excel.Application _application;

            internal ExcelController(BannerPaneManager panes, Excel.Application application)
                : base(panes, application, "Excel")
            {
                _application = application;
            }

            protected override object GetActiveDocument()
            {
                return _application.ActiveWorkbook;
            }

            protected override object GetWindowFor(object document)
            {
                var workbook = (Excel.Workbook)document;
                return workbook.Windows.Count == 0 ? null : workbook.Windows[1];
            }

            protected override DocumentMacroInfo Describe(object document)
            {
                var workbook = (Excel.Workbook)document;

                return new DocumentMacroInfo
                {
                    HasVbaProject = workbook.HasVBProject,
                    IsVbaSigned = workbook.HasVBProject && workbook.VBASigned,
                    HasExcel4Macros = HasExcel4MacroSheets(workbook),
                    FullPath = GetFullPath(workbook)
                };
            }

            /// <summary>
            /// Excel 4.0 macro sheets are macros that <c>HasVBProject</c> does not report. They
            /// predate VBA, cannot be signed, and have a long history as a malware vector, so a
            /// tool that answers "does this file have a macro?" has to look for them.
            /// </summary>
            private static bool HasExcel4MacroSheets(Excel.Workbook workbook)
            {
                try
                {
                    return workbook.Excel4MacroSheets.Count > 0
                        || workbook.Excel4IntlMacroSheets.Count > 0;
                }
                catch (Exception ex)
                {
                    Log.Warn("Could not enumerate Excel 4.0 macro sheets", ex);
                    return false;
                }
            }

            /// <summary>
            /// A workbook that has never been saved has no path, and Excel reports SharePoint and
            /// OneDrive workbooks with a URL. Neither can be checked against a Trusted Location or
            /// a zone identifier, so both come back empty.
            /// </summary>
            private static string GetFullPath(Excel.Workbook workbook)
            {
                try
                {
                    return string.IsNullOrEmpty(workbook.Path) ? null : workbook.FullName;
                }
                catch (Exception ex)
                {
                    Log.Warn("Could not read the workbook path", ex);
                    return null;
                }
            }

            protected override void ShowNoDocumentMessage()
            {
                MessageBox.Show(
                    Strings.Get("Error_NoDocument"),
                    Strings.Get("Ribbon_GroupLabel"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }

            protected override void ShowCheckFailedMessage(Exception exception)
            {
                MessageBox.Show(
                    Strings.Get("Error_CheckFailed") + Environment.NewLine + Environment.NewLine + exception.Message,
                    Strings.Get("Ribbon_GroupLabel"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }
    }
}
