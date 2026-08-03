using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Macro_Polo.Core;
using Macro_Polo_Word.Properties;
using Word = Microsoft.Office.Interop.Word;

namespace Macro_Polo_Word
{
    /// <summary>
    /// The Word add-in. Office creates this class by ProgID and drives it through
    /// <see cref="IDTExtensibility2"/>.
    /// </summary>
    /// <remarks>
    /// The class id and ProgID are fixed and must not change once shipped: the installer writes
    /// them into the registry, and an installed machine would otherwise be left pointing at a
    /// class that no longer exists. The ProgID doubles as the key name Office looks for under
    /// <c>Software\Microsoft\Office\Word\Addins</c>.
    /// </remarks>
    [ComVisible(true)]
    [Guid(Clsid)]
    [ProgId(ProgId)]
    // Office resolves ribbon callbacks by name through IDispatch, which AutoDispatch provides by
    // reflecting over the class at run time. AutoDual would additionally bake a static dual vtable
    // out of every public member - including ones typed in embedded interop types - which buys
    // nothing here and is fragile across rebuilds.
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class WordAddIn : ComAddInBase
    {
        /// <summary>Class id registered for this add-in. Referenced by the installer.</summary>
        public const string Clsid = "1B9E6F42-7A35-4C80-9D13-2E5C4A8B7061";

        /// <summary>ProgID, and the name of this add-in's key under Office's Addins key.</summary>
        public const string ProgId = "Macro_Polo.WordAddIn";

        private Word.Application _application;

        protected override string RibbonXmlResourceName
        {
            get { return "Macro_Polo_Word.Ribbon1.xml"; }
        }

        protected override Bitmap ButtonBitmap
        {
            get { return Resources.word; }
        }

        protected override MacroPoloController CreateController(object application, BannerPaneManager panes)
        {
            _application = application as Word.Application;

            if (_application == null)
            {
                throw new InvalidOperationException("Macro Polo was loaded by a host that is not Word.");
            }

            return new WordController(panes, _application);
        }

        protected override void ConnectHostEvents()
        {
            _application.DocumentOpen += OnDocumentPresented;
            _application.WindowActivate += OnWindowActivate;
            _application.DocumentBeforeClose += OnDocumentBeforeClose;
        }

        protected override void DisconnectHostEvents()
        {
            if (_application == null)
            {
                return;
            }

            _application.DocumentOpen -= OnDocumentPresented;
            _application.WindowActivate -= OnWindowActivate;
            _application.DocumentBeforeClose -= OnDocumentBeforeClose;
            _application = null;
        }

        private void OnDocumentPresented(Word.Document document)
        {
            Controller.HandleDocumentPresented(document);
        }

        /// <summary>
        /// Covers documents that were already open when the add-in loaded, documents created from
        /// a template, and switching between open documents.
        /// </summary>
        private void OnWindowActivate(Word.Document document, Word.Window window)
        {
            Controller.HandleDocumentPresented(document);
        }

        private void OnDocumentBeforeClose(Word.Document document, ref bool cancel)
        {
            Controller.HandleDocumentClosed(document);
        }

        /// <summary>Word's answers to the host-specific questions.</summary>
        private sealed class WordController : MacroPoloController
        {
            private readonly Word.Application _application;

            internal WordController(BannerPaneManager panes, Word.Application application)
                : base(panes, application, "Word")
            {
                _application = application;
            }

            protected override object GetActiveDocument()
            {
                // Word throws rather than returning null when nothing is open.
                return _application.Documents.Count == 0 ? null : _application.ActiveDocument;
            }

            protected override object GetWindowFor(object document)
            {
                var doc = (Word.Document)document;
                return doc.Windows.Count == 0 ? null : doc.ActiveWindow;
            }

            protected override DocumentMacroInfo Describe(object document)
            {
                var doc = (Word.Document)document;

                return new DocumentMacroInfo
                {
                    HasVbaProject = doc.HasVBProject,
                    IsVbaSigned = doc.HasVBProject && doc.VBASigned,
                    FullPath = GetFullPath(doc)
                };
            }

            /// <summary>
            /// A document that has never been saved has no path, and Word reports SharePoint and
            /// OneDrive documents with a URL. Neither can be checked against a Trusted Location or
            /// a zone identifier, so both come back empty.
            /// </summary>
            private static string GetFullPath(Word.Document document)
            {
                try
                {
                    return string.IsNullOrEmpty(document.Path) ? null : document.FullName;
                }
                catch (Exception ex)
                {
                    Log.Warn("Could not read the document path", ex);
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
