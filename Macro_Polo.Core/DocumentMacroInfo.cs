namespace Macro_Polo.Core
{
    /// <summary>
    /// Everything the evaluator needs to know about a single document. Populated by the host
    /// add-in from the Word or Excel object model so that <see cref="MacroStatusEvaluator"/>
    /// stays free of any Office dependency.
    /// </summary>
    public sealed class DocumentMacroInfo
    {
        /// <summary>The document contains a VBA project.</summary>
        public bool HasVbaProject { get; set; }

        /// <summary>
        /// The workbook contains legacy Excel 4.0 (XLM) macro sheets. These are macros that
        /// <c>HasVBProject</c> does not report, and they are governed by the same macro settings.
        /// Always false for Word.
        /// </summary>
        public bool HasExcel4Macros { get; set; }

        /// <summary>
        /// The VBA project carries a digital signature. Note that Office exposes this as a bare
        /// flag: it tells us a signature blob is present, not that the certificate is valid,
        /// unexpired, or trusted. See <see cref="MacroState.RequiresPublisherTrust"/>.
        /// </summary>
        public bool IsVbaSigned { get; set; }

        /// <summary>Full path of the file on disk, or null/empty for a document that was never saved.</summary>
        public string FullPath { get; set; }

        /// <summary>True when the file carries a mark of the web identifying it as internet or restricted zone.</summary>
        public bool HasMarkOfTheWeb { get; set; }

        /// <summary>True when any macro code is present, from either VBA or Excel 4.0 macro sheets.</summary>
        public bool HasMacros
        {
            get { return HasVbaProject || HasExcel4Macros; }
        }
    }
}
