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
        Normal = 0b_0000_0000,

        /// <summary>
        /// A bold text
        /// </summary>
        Bold = 0b_0000_0001,

        /// <summary>
        /// A italic text
        /// </summary>
        Italic = 0b_0000_0010,

        /// <summary>
        /// A text with a underline
        /// </summary>
        Underline = 0x_0000_0100,

        /// <summary>
        /// The text alignment is left
        /// </summary>
        Left = 0x_0001_0000,

        /// <summary>
        /// The text alignment is centered
        /// </summary>
        Center = 0x_0010_0000,

        /// <summary>
        /// The text alignment is right
        /// </summary>
        Right = 0x_0100_0000,

        /// <summary>
        /// The text use full justification
        /// </summary>
        Justify = 0x_1000_0000,
    }
}
