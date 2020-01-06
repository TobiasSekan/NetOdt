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
        None = 0,

        /// <summary>
        /// A bold text
        /// </summary>
        Bold = 1 << 0,

        /// <summary>
        /// A italic text
        /// </summary>
        Italic = 1 << 1,

        /// <summary>
        /// A text with a single underline
        /// </summary>
        UnderlineSingle = 1 << 2,

        /// <summary>
        /// A text with a double underline
        /// </summary>
        UnderlineDouble = 1 << 3,

        /// <summary>
        /// A text with a bold underline
        /// </summary>
        UnderlineBold = 1 << 4,

        /// <summary>
        /// A text with a dotted underline
        /// </summary>
        UnderlineDotted = 1 << 5,

        /// <summary>
        /// A text with a bold dotted underline
        /// </summary>
        UnderlineDottedBold = 1 << 6,

        /// <summary>
        /// A text with a dash underline
        /// </summary>
        UnderlineDash = 1 << 7,

        /// <summary>
        /// A text with a long-dash underline
        /// </summary>
        UnderlineLongDash = 1 << 8,

        /// <summary>
        /// A text with a dot-dash underline
        /// </summary>
        UnderlineDotDash = 1 << 9,

        /// <summary>
        /// A text with a dot-dot-dash underline
        /// </summary>
        UnderlineDotDotDash = 1 << 10,

        /// <summary>
        /// A text with a wave underline
        /// </summary>
        UnderlineWave = 1 << 11,

        /// <summary>
        /// A stroked text
        /// </summary>
        Stroke = 1 << 12,

        /// <summary>
        /// A subscript text
        /// </summary>
        Subscript = 1 << 13,

        /// <summary>
        /// A superscript text
        /// </summary>
        Superscript = 1 << 14,

        /// <summary>
        /// The text alignment is left
        /// </summary>
        Left = 1 << 15,

        /// <summary>
        /// The text alignment is centered
        /// </summary>
        Center = 1 << 16,

        /// <summary>
        /// The text alignment is right
        /// </summary>
        Right = 1 << 17,

        /// <summary>
        /// The text use full justification
        /// </summary>
        Justify = 1 << 18,
    }
}
