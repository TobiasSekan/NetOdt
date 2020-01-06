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
        Normal = 0,

        /// <summary>
        /// A bold text
        /// </summary>
        Bold = 1 << 0,

        /// <summary>
        /// A italic text
        /// </summary>
        Italic = 1 << 1,

        /// <summary>
        /// A text with a underline (single stroke)
        /// </summary>
        UnderlineSingle = 1 << 2,

        /// <summary>
        /// A stroked text
        /// </summary>
        Stroke = 1 << 3,

        /// <summary>
        /// A subscript text
        /// </summary>
        Subscript = 1 << 4,

        /// <summary>
        /// A superscript text
        /// </summary>
        Superscript = 1 << 5,

        /// <summary>
        /// The text alignment is left
        /// </summary>
        Left = 1 << 6,

        /// <summary>
        /// The text alignment is centered
        /// </summary>
        Center = 1 << 7,

        /// <summary>
        /// The text alignment is right
        /// </summary>
        Right = 1 << 8,

        /// <summary>
        /// The text use full justification
        /// </summary>
        Justify = 1 << 9,
    }
}
