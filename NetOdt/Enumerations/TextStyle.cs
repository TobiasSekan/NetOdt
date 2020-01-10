using System;

namespace NetOdt.Enumerations
{
    /// <summary>
    /// The style of a text
    /// </summary>
    [Flags]
    public enum TextStyle : ulong
    {
        /// <summary>
        /// Text without any style
        /// </summary>
        None = 0x_00_00_00_00_00_00,

        /// <summary>
        /// A bold text
        /// </summary>
        Bold = 0x_00_00_00_00_00_01,

        /// <summary>
        /// A italic text
        /// </summary>
        Italic = 0x_00_00_00_00_00_02,

        /// <summary>
        /// A text with a single underline
        /// </summary>
        UnderlineSingle = 0x_00_00_00_00_00_04,

        /// <summary>
        /// A text with a double underline
        /// </summary>
        UnderlineDouble = 0x_00_00_00_00_00_08,

        /// <summary>
        /// A text with a bold underline
        /// </summary>
        UnderlineBold = 0x_00_00_00_00_00_10,

        /// <summary>
        /// A text with a dotted underline
        /// </summary>
        UnderlineDotted = 0x_00_00_00_00_00_20,

        /// <summary>
        /// A text with a bold dotted underline
        /// </summary>
        UnderlineDottedBold = 0x_00_00_00_00_00_40,

        /// <summary>
        /// A text with a dash underline
        /// </summary>
        UnderlineDash = 0x_00_00_00_00_00_80,

        /// <summary>
        /// A text with a long-dash underline
        /// </summary>
        UnderlineLongDash = 0x_00_00_00_00_01_00,

        /// <summary>
        /// A text with a dot-dash underline
        /// </summary>
        UnderlineDotDash = 0x_00_00_00_00_02_00,

        /// <summary>
        /// A text with a dot-dot-dash underline
        /// </summary>
        UnderlineDotDotDash = 0x_00_00_00_00_04_00,

        /// <summary>
        /// A text with a wave underline
        /// </summary>
        UnderlineWave = 0x_00_00_00_00_08_00,

        /// <summary>
        /// A stroked text
        /// </summary>
        Stroke = 0x_00_00_00_00_10_00,

        /// <summary>
        /// A subscript text
        /// </summary>
        Subscript = 0x_00_00_00_00_20_00,

        /// <summary>
        /// A superscript text
        /// </summary>
        Superscript = 0x_00_00_00_00_40_00,

        /// <summary>
        /// The text alignment is left
        /// </summary>
        Left = 0x_00_00_00_00_80_00,

        /// <summary>
        /// The text alignment is centered
        /// </summary>
        Center = 0x_00_00_00_01_00_00,

        /// <summary>
        /// The text alignment is right
        /// </summary>
        Right = 0x_00_00_00_02_00_00,

        /// <summary>
        /// The text use full justification
        /// </summary>
        Justify = 0x_00_00_00_04_00_00,

        /// <summary>
        /// Title
        /// <para>font size 28 and bold and centered</para>
        /// </summary>
        Title =  0x_00_00_00_08_00_00,

        /// <summary>
        /// Subtitle
        /// <para>font size 18 and centered</para>
        /// </summary>
        Subtitle = 0x_00_00_00_10_00_00,

        /// <summary>
        /// Heading level 1
        /// <para>font size 18,2 and bold</para>
        /// </summary>
        HeadingLevel01 = 0x_00_00_00_20_00_00,

        /// <summary>
        /// Heading level 2
        /// <para>font size 16,1 and bold</para>
        /// </summary>
        HeadingLevel02 = 0x_00_00_00_40_00_00,

        /// <summary>
        /// Heading level 3
        /// <para>font size 14,1 and bold</para>
        /// </summary>
        HeadingLevel03 = 0x_00_00_00_80_00_00,

        /// <summary>
        /// Heading level 4
        /// <para>font size 13,3 and bold and italic</para>
        /// </summary>
        HeadingLevel04 = 0x_00_00_01_00_00_00,

        /// <summary>
        /// Heading level 5
        /// <para>font size 11,8 and bold</para>
        /// </summary>
        HeadingLevel05 = 0x_00_00_02_00_00_00,

        /// <summary>
        /// Heading level 6
        /// <para>font size 11,8 and bold and italic</para>
        /// </summary>
        HeadingLevel06 = 0x_00_00_04_00_00_00,

        /// <summary>
        /// Heading level 7
        /// <para>font size 11,1 and bold</para>
        /// </summary>
        HeadingLevel07 = 0x_00_00_08_00_00_00,

        /// <summary>
        /// Heading level 8
        /// <para>font size 11,1 and bold and italic</para>
        /// </summary>
        HeadingLevel08 = 0x_00_00_10_00_00_00,

        /// <summary>
        /// Heading level 9
        /// <para>font size 10,5 and bold</para>
        /// </summary>
        HeadingLevel09 = 0x_00_00_20_00_00_00,

        /// <summary>
        /// Heading level 10
        /// <para>font size 10,5 and bold</para>
        /// </summary>
        HeadingLevel10 = 0x_00_00_40_00_00_00,

        /// <summary>
        /// Signature
        /// </summary>
        Signature = 0x_00_00_80_00_00_00,

        /// <summary>
        /// Quotations
        /// </summary>
        Quotations = 0x_00_01_00_00_00_00,

        /// <summary>
        /// Endnote
        /// </summary>
        Endnote = 0x_00_02_00_00_00_00,

        /// <summary>
        /// Footnote
        /// </summary>
        Footnote = 0x_00_04_00_00_00_00,

        /// <summary>
        /// Page break before text
        /// </summary>
        PageBreak = 0x_00_08_00_00_00_00,
    }
}
