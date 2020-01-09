using System;

namespace NetOdt.Enumerations
{
    /// <summary>
    /// The style of a text
    /// </summary>
    [Flags]
    public enum TextStyle : long
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

        /// <summary>
        /// Title
        /// <para>font size 28 and bold and centered</para>
        /// </summary>
        Title =  1 << 19,

        /// <summary>
        /// Subtitle
        /// <para>font size 18 and centered</para>
        /// </summary>
        Subtitle = 1 << 20,

        /// <summary>
        /// Heading level 1
        /// <para>font size 18,2 and bold</para>
        /// </summary>
        HeadingLevel01 = 1 << 21,

        /// <summary>
        /// Heading level 2
        /// <para>font size 16,1 and bold</para>
        /// </summary>
        HeadingLevel02 = 1 << 22,

        /// <summary>
        /// Heading level 3
        /// <para>font size 14,1 and bold</para>
        /// </summary>
        HeadingLevel03 = 1 << 23,

        /// <summary>
        /// Heading level 4 ()
        /// <para>font size 13,3 and bold and italic</para>
        /// </summary>
        HeadingLevel04 = 1 << 24,

        /// <summary>
        /// Heading level 5 ()
        /// <para>font size 11,8 and bold</para>
        /// </summary>
        HeadingLevel05 = 1 << 25,

        /// <summary>
        /// Heading level 6
        /// <para>font size 11,8 and bold and italic</para>
        /// </summary>
        HeadingLevel06 = 1 << 26,

        /// <summary>
        /// Heading level 7
        /// <para>font size 11,1 and bold</para>
        /// </summary>
        HeadingLevel07 = 1 << 27,

        /// <summary>
        /// Heading level 8
        /// <para>font size 11,1 and bold and italic</para>
        /// </summary>
        HeadingLevel08 = 1 << 28,

        /// <summary>
        /// Heading level 9
        /// <para>font size 10,5 and bold</para>
        /// </summary>
        HeadingLevel09 = 1 << 29,

        /// <summary>
        /// Heading level 10
        /// <para>font size 10,5 and bold</para>
        /// </summary>
        HeadingLevel10 = 1 << 30,

        /// <summary>
        /// Signature
        /// </summary>
        Signature = 1 << 31,

        /// <summary>
        /// Quotations
        /// </summary>
        Quotations = 1 << 32,

        /// <summary>
        /// Endnote
        /// </summary>
        Endnote = 1 << 33,

        /// <summary>
        /// Footnote
        /// </summary>
        Footnote = 1 << 34,
    }
}
