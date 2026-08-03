using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

// The trust decision in VbaSignatureReader is worth testing directly, but not worth exposing:
// callers should go through Read(path), which is the only entry point that makes sense.
[assembly: InternalsVisibleTo("Macro_Polo.Core.Tests")]

// The SDK's generated assembly info does not emit ComVisible, and the default is visible. Left
// alone, every public type in this library gets its own CLSID and ProgID, and regasm - or the
// installer, which mirrors it - would register a dozen classes that have no business being
// creatable from COM. Only MacroStatusBanner opts back in, because Office's task pane factory has
// to be able to create it by ProgID.
[assembly: ComVisible(false)]
