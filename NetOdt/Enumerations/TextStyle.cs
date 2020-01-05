using System;

namespace NetOdt.Enumerations
{
    /// <summary>
    /// The style of a text
    /// </summary>
    [Flags]
    public enum TextStyle
    {
        /// <summary>
        /// Text without any style
        /// </summary>
        Normal = 0x_00,

        /// <summary>
        /// A bold text
        /// </summary>
        Bold = 0x_01,

        /// <summary>
        /// A italic text
        /// </summary>
        Italic = 0x_02,

        /// <summary>
        /// A text with a underline
        /// </summary>
        Underline = 0x_04,
    }
}
