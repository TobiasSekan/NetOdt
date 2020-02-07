using NetOdt.Enumerations;
using System;
using System.Drawing;

namespace NetOdt.Class
{
    /// <summary>
    /// A style for a element passage
    /// </summary>
    internal sealed class Style
    {
        #region Internal Properties

        /// <summary>
        /// The name of the font of this style
        /// </summary>
        internal string FontName { get; }

        /// <summary>
        /// The size of the font of this style
        /// </summary>
        internal double FontSize { get; }

        /// <summary>
        /// The name of the style
        /// </summary>
        internal string Name { get; }

        /// <summary>
        /// The name of the style family
        /// </summary>
        internal string FamilyName { get; }

        /// <summary>
        /// The name of the parent style (inheritance style informations)
        /// </summary>
        internal string ParentName { get; }

        /// <summary>
        /// The type of the style family
        /// </summary>
        internal StyleFamily StyleFamily { get; }

        /// <summary>
        /// The type of the text style
        /// </summary>
        internal TextStyle TextStyle { get; }

        /// <summary>
        /// The outline level of this style (only used by headings)
        /// </summary>
        internal byte OutlineLevel { get; }

        /// <summary>
        /// Indicate the style have a parent font size
        /// </summary>
        internal bool HaveParentFontSize { get; }

        /// <summary>
        /// The <see cref="Color"/> of the foreground
        /// </summary>
        internal Color ForegroundColor { get; }

        /// <summary>
        /// The <see cref="Color"/> of the background
        /// </summary>
        internal Color BackgroundColor { get; }

        /// <summary>
        /// The "office value" type for this style
        /// </summary>
        internal OfficeValueType ValueType { get; }

        #endregion Internal Properties

        #region Internal Constructors

        /// <summary>
        /// Create a new style with the given informations
        /// </summary>
        /// <param name="name">The name for this style</param>
        /// <param name="textStyle">The <see cref="TextStyle"/> for this style</param>
        /// <param name="styleFamily">The <see cref="StyleFamily"/> for this style</param>
        /// <param name="valueType">The <see cref="OfficeValueType"/> for this style</param>
        /// <param name="fontSize">The size of the font for this style</param>
        /// <param name="fontName">The name of the font for this style</param>
        /// <param name="foreground">The foreground color for this style</param>
        /// <param name="background">The background color for this style</param>
        internal Style(string name,
                       TextStyle textStyle,
                       StyleFamily styleFamily,
                       OfficeValueType valueType,
                       double fontSize,
                       string fontName,
                       Color foreground,
                       Color background)
        {
            FontName           = fontName;
            FontSize           = fontSize;
            Name               = name;
            StyleFamily        = styleFamily;
            TextStyle          = textStyle;
            ForegroundColor    = foreground;
            BackgroundColor    = background;
            ValueType          = valueType;

            FamilyName         = GetStyleFamily();
            ParentName         = GetParentStyleName();
            HaveParentFontSize = CheckParentFontSize();
            OutlineLevel       = GetOutlineLevel();
        }

        /// <summary>
        /// Create a new style with the given informations
        /// </summary>
        /// <param name="name">The name for this style</param>
        /// <param name="family">The style family for this style</param>
        /// <param name="valueType">The type of the "office value" for this style</param>
        internal Style(string name, StyleFamily family, OfficeValueType valueType)
        {
            FontName           = string.Empty;
            FontSize           = 0.0;
            Name               = name;
            StyleFamily             = family;
            TextStyle          = TextStyle.None;
            ForegroundColor    = Color.Black;
            BackgroundColor    = Color.Transparent;
            ValueType          = valueType;

            FamilyName         = GetStyleFamily();
            ParentName         = GetParentStyleName();
            HaveParentFontSize = CheckParentFontSize();
            OutlineLevel       = GetOutlineLevel();
        }

        #endregion Internal Constructors

        #region Internal Methods

        /// <summary>
        /// Return the style family name of this style
        /// </summary>
        /// <returns>A style family name</returns>
        /// <exception cref="ArgumentOutOfRangeException">"Style family not supported"</exception>
        internal string GetStyleFamily()
            => StyleFamily switch
            {
                // TODO: don't use StyleFamily for header and footer
                StyleFamily.Header      => "paragraph",
                StyleFamily.Footer      => "paragraph",

                StyleFamily.Paragraph   => "paragraph",
                StyleFamily.Table       => "table",
                StyleFamily.TableColumn => "table-column",
                StyleFamily.TableCell   => "table-cell",
                StyleFamily.Graphic     => "graphic",

                _                       => throw new ArgumentOutOfRangeException(nameof(StyleFamily),
                                                                                 StyleFamily,
                                                                                 "Style family not supported"),
            };

        /// <summary>
        /// Return the parent style name of this style
        /// </summary>
        /// <returns>A parent style name</returns>
        internal string GetParentStyleName()
        {
            switch(StyleFamily)
            {
                case StyleFamily.Graphic:
                    return "Graphics";

                case StyleFamily.Header:
                    return "Header";

                case StyleFamily.Footer:
                    return "Footer";

                case StyleFamily.Paragraph:
                    break;

                default:
                    return string.Empty;
            }

            return TextStyle switch
            {
                TextStyle.HeadingLevel01 => "Heading_20_1",
                TextStyle.HeadingLevel02 => "Heading_20_2",
                TextStyle.HeadingLevel03 => "Heading_20_3",
                TextStyle.HeadingLevel04 => "Heading_20_4",
                TextStyle.HeadingLevel05 => "Heading_20_5",
                TextStyle.HeadingLevel06 => "Heading_20_6",
                TextStyle.HeadingLevel07 => "Heading_20_7",
                TextStyle.HeadingLevel08 => "Heading_20_8",
                TextStyle.HeadingLevel09 => "Heading_20_9",
                TextStyle.HeadingLevel10 => "Heading_20_10",

                TextStyle.Title          => "Title",
                TextStyle.Subtitle       => "Subtitle",
                TextStyle.Signature      => "Signature",
                TextStyle.Quotations     => "Quotations",
                TextStyle.Endnote        => "Endnote",
                TextStyle.Footnote       => "Footnote",

                _                        => "Standard",
            };
        }

        /// <summary>
        /// Check if the text style of this style have a parent font size
        /// </summary>
        /// <returns><see langword="true"/> if the given text style have a parent font size, otherwise <see langword="false"/></returns>
        internal bool CheckParentFontSize()
            => TextStyle switch
            {
                TextStyle.HeadingLevel01 => true,
                TextStyle.HeadingLevel02 => true,
                TextStyle.HeadingLevel03 => true,
                TextStyle.HeadingLevel04 => true,
                TextStyle.HeadingLevel05 => true,
                TextStyle.HeadingLevel06 => true,
                TextStyle.HeadingLevel07 => true,
                TextStyle.HeadingLevel08 => true,
                TextStyle.HeadingLevel09 => true,
                TextStyle.HeadingLevel10 => true,

                TextStyle.Title          => true,
                TextStyle.Subtitle       => true,
                TextStyle.Signature      => true,
                TextStyle.Quotations     => true,
                TextStyle.Endnote        => true,
                TextStyle.Footnote       => true,

                _                        => false,
            };

        /// <summary>
        /// Return the outline level of the <see cref="TextStyle"/>
        /// </summary>
        /// <returns>A outline level</returns>
        internal byte GetOutlineLevel()
            => TextStyle switch
            {
                TextStyle.HeadingLevel01 => 01,
                TextStyle.HeadingLevel02 => 02,
                TextStyle.HeadingLevel03 => 03,
                TextStyle.HeadingLevel04 => 04,
                TextStyle.HeadingLevel05 => 05,
                TextStyle.HeadingLevel06 => 06,
                TextStyle.HeadingLevel07 => 07,
                TextStyle.HeadingLevel08 => 08,
                TextStyle.HeadingLevel09 => 09,
                TextStyle.HeadingLevel10 => 10,
                _                        => 00,
            };

        #endregion Internal Methods
    }
}
