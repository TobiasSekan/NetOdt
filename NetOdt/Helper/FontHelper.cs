using NetOdt.Enumerations;

namespace NetOdt.Helper
{
    /// <summary>
    /// Helper class to easier work with fonts
    /// </summary>
    internal static class FontHelper
    {
        /// <summary>
        /// Return the real size of a give font size representation for a font
        /// </summary>
        /// <param name="fontSize">The font size representation for the real font size</param>
        /// <returns>A real font size</returns>
        internal static double GetFontSize(FontSize fontSize)
            => fontSize == FontSize.Size10Point5 ? 10.5 : (double)fontSize;
    }
}
