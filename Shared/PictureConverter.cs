using System.Drawing;
using System.Windows.Forms;

namespace Macro_Polo.Core
{
    /// <summary>
    /// Converts a <see cref="Image"/> into the <c>IPictureDisp</c> the ribbon's
    /// <c>getImage</c> callback has to return.
    /// </summary>
    /// <remarks>
    /// VSTO used to accept a plain <see cref="Bitmap"/> from the callback and do this conversion
    /// on the way out. A raw COM add-in talks to Office directly, so the conversion is ours to do.
    /// <see cref="AxHost"/> holds the only accessible implementation of it in the framework, and
    /// its converter is protected, hence the otherwise pointless subclass.
    /// </remarks>
    internal sealed class PictureConverter : AxHost
    {
        private PictureConverter()
            : base(string.Empty)
        {
        }

        internal static stdole.IPictureDisp ToPictureDisp(Image image)
        {
            return image == null ? null : (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
        }
    }
}
