using NetOdt.Enumerations;
using System;
using System.Text;

namespace NetOdt.Helper
{
    /// <summary>
    /// Helper class to easier work with ODT document styles
    /// </summary>
    internal static class StyleHelper
    {
        /// <summary>
        /// Add all needed styles for all <see cref="TextStyle"/> combinations to the style content
        /// </summary>
        /// <param name="styleContent">The content container for the style informations</param>
        internal static void AddStandardTextStyles(in StringBuilder styleContent)
        {
            AppendXmlStyleStart(styleContent, "P1", StyleFamily.Paragraph);
            AppendXmlStyle(styleContent, TextStyle.Normal);
            AppendXmlStyleEnd(styleContent);

            AppendXmlStyleStart(styleContent, "P2", StyleFamily.Paragraph);
            AppendXmlStyle(styleContent, TextStyle.Bold);
            AppendXmlStyleEnd(styleContent);

            // P3 - 0b_010 - Italic
            AppendXmlStyleStart(styleContent, "P3", StyleFamily.Paragraph);
            AppendXmlStyle(styleContent, TextStyle.Italic);
            AppendXmlStyleEnd(styleContent);

            // P4 - 0b_011 - Bold + Italic
            AppendXmlStyleStart(styleContent, "P4", StyleFamily.Paragraph);
            AppendXmlStyle(styleContent, TextStyle.Bold | TextStyle.Italic);
            AppendXmlStyleEnd(styleContent);

            // P5 0b_100 - Underline
            AppendXmlStyleStart(styleContent, "P5", StyleFamily.Paragraph);
            AppendXmlStyle(styleContent, TextStyle.Underline);
            AppendXmlStyleEnd(styleContent);

            // P6 - 0b_101 - Bold + Underline
            AppendXmlStyleStart(styleContent, "P6", StyleFamily.Paragraph);
            AppendXmlStyle(styleContent, TextStyle.Bold | TextStyle.Underline);
            AppendXmlStyleEnd(styleContent);

            // P7 - 0b_110 - Italic + Underline
            AppendXmlStyleStart(styleContent, "P7", StyleFamily.Paragraph);
            AppendXmlStyle(styleContent, TextStyle.Italic | TextStyle.Underline);
            AppendXmlStyleEnd(styleContent);

            // P8 - 0b_111 - Bold + Italic + Underline
            AppendXmlStyleStart(styleContent, "P8", StyleFamily.Paragraph);
            AppendXmlStyle(styleContent, TextStyle.Bold | TextStyle.Italic | TextStyle.Underline);
            AppendXmlStyleEnd(styleContent);
        }

        /// <summary>
        /// Add all needed styles for all <see cref="TextAlignment"/> combinations to the style content
        /// </summary>
        /// <param name="styleContent">The content container for the style informations</param>
        internal static void AddTextAlignmentStyles(in StringBuilder styleContent)
        {
            // Left text alignment
            AppendXmlStyleStart(styleContent, "P21", StyleFamily.Paragraph);
            styleContent.Append("<style:paragraph-properties fo:text-align=\"start\" style:justify-single-word=\"false\"/>");
            AppendXmlStyleEnd(styleContent);

            // Centered text
            AppendXmlStyleStart(styleContent, "P22", StyleFamily.Paragraph);
            styleContent.Append("<style:paragraph-properties fo:text-align=\"center\" style:justify-single-word=\"false\"/>");
            AppendXmlStyleEnd(styleContent);

            // Right text alignment
            AppendXmlStyleStart(styleContent, "P23", StyleFamily.Paragraph);
            styleContent.Append("<style:paragraph-properties fo:text-align=\"end\" style:justify-single-word=\"false\"/>");
            AppendXmlStyleEnd(styleContent);

            // Full justification
            AppendXmlStyleStart(styleContent, "P24", StyleFamily.Paragraph);
            styleContent.Append("<style:paragraph-properties fo:text-align=\"justify\" style:justify-single-word=\"false\"/>");
            AppendXmlStyleEnd(styleContent);
        }

        /// <summary>
        /// Add all needed styles for simple tables
        /// </summary>
        /// <param name="styleContent">The content container for the style informations</param>
        internal static void AddTableStyles(in StringBuilder styleContent)
        {
            AppendXmlStyleStart(styleContent, "Tabelle1", StyleFamily.Table);
            styleContent.Append("<style:table-properties style:width=\"17cm\" table:align=\"margins\"/>");
            AppendXmlStyleEnd(styleContent);

            AppendXmlStyleStart(styleContent, "Tabelle1.A", StyleFamily.TableColumn);
            styleContent.Append("<style:table-column-properties style:column-width=\"3.401cm\" style:rel-column-width=\"13107*\"/>");
            AppendXmlStyleEnd(styleContent);

            AppendXmlStyleStart(styleContent, "Tabelle1.A1", StyleFamily.TableCell);
            styleContent.Append("<style:table-cell-properties fo:padding=\"0.097cm\" fo:border-left=\"0.05pt solid #000000\" fo:border-right=\"none\" fo:border-top=\"0.05pt solid #000000\" fo:border-bottom=\"0.05pt solid #000000\"/>");
            AppendXmlStyleEnd(styleContent);

            AppendXmlStyleStart(styleContent, "Tabelle1.E1", StyleFamily.TableCell);
            styleContent.Append("<style:table-cell-properties fo:padding=\"0.097cm\" fo:border=\"0.05pt solid #000000\"/>");
            AppendXmlStyleEnd(styleContent);

            AppendXmlStyleStart(styleContent, "Tabelle1.A2", StyleFamily.TableCell);
            styleContent.Append("<style:table-cell-properties fo:padding=\"0.097cm\" fo:border-left=\"0.05pt solid #000000\" fo:border-right=\"none\" fo:border-top=\"none\" fo:border-bottom=\"0.05pt solid #000000\"/>");
            AppendXmlStyleEnd(styleContent);

            AppendXmlStyleStart(styleContent, "Tabelle1.E2", StyleFamily.TableCell);
            styleContent.Append("<style:table-cell-properties fo:padding=\"0.097cm\" fo:border-left=\"0.05pt solid #000000\" fo:border-right=\"0.05pt solid #000000\" fo:border-top=\"none\" fo:border-bottom=\"0.05pt solid #000000\"/>");
            AppendXmlStyleEnd(styleContent);
        }

        /// <summary>
        /// Add all needed styles for simple pictures
        /// </summary>
        /// <param name="styleContent">The content container for the style informations</param>
        internal static void AddPictureStyle(in StringBuilder styleContent)
        {
            AppendXmlStyleStart(styleContent, "fr1", StyleFamily.Graphic);
            styleContent.Append("<style:graphic-properties style:mirror=\"none\" fo:clip=\"rect(0cm, 0cm, 0cm, 0cm)\" draw:luminance=\"0%\" draw:contrast=\"0%\" draw:red=\"0%\" draw:green=\"0%\" draw:blue=\"0%\" draw:gamma=\"100%\" draw:color-inversion=\"false\" draw:image-opacity=\"100%\" draw:color-mode=\"standard\"/>");
            AppendXmlStyleEnd(styleContent);
        }

        /// <summary>
        /// Append a XML start element for style informations
        /// </summary>
        /// <param name="styleContent">The content container for the style informations</param>
        /// <param name="styleName">The name for the style</param>
        /// <param name="styleFamily">The family of the style</param>
        /// <exception cref="ArgumentOutOfRangeException">Style family not supported</exception>
        internal static void AppendXmlStyleStart(in StringBuilder styleContent, in string styleName, in StyleFamily styleFamily)
        {
            switch(styleFamily)
            {
                case StyleFamily.Paragraph:
                    styleContent.Append($"<style:style style:name=\"{styleName}\" style:family=\"paragraph\" style:parent-style-name=\"Standard\">");
                    break;

                case StyleFamily.Table:
                    styleContent.Append($"<style:style style:name=\"{styleName}\" style:family=\"table\">");
                    break;

                case StyleFamily.TableColumn:
                    styleContent.Append($"<style:style style:name=\"{styleName}\" style:family=\"table-column\">");
                    break;

                case StyleFamily.TableCell:
                    styleContent.Append($"<style:style style:name=\"{styleName}\" style:family=\"table-cell\">");
                    break;

                case StyleFamily.Graphic:
                    styleContent.Append($"<style:style style:name=\"{styleName}\" style:family=\"graphic\" style:parent-style-name=\"Graphics\">");
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(styleFamily), styleFamily, "Style family not supported");
            }
        }

        /// <summary>
        /// Append the given text style to the style informations
        /// </summary>
        /// <param name="styleContent">The content container for the style informations</param>
        /// <param name="style">The style for the style informations</param>
        internal static void AppendXmlStyle(in StringBuilder styleContent, TextStyle style)
        {
            // Note: Don't forget the spaces between the XML properties

            switch(style)
            {
                case TextStyle.Normal:
                    styleContent.Append("<style:text-properties/>");
                    break;

                case TextStyle.Bold:
                    styleContent.Append("<style:text-properties fo:font-weight=\"bold\" style:font-weight-asian=\"bold\" style:font-weight-complex=\"bold\"/>");
                    break;

                case TextStyle.Italic:
                    styleContent.Append("<style:text-properties fo:font-style=\"italic\" style:font-style-asian=\"italic\" style:font-style-complex=\"italic\"/>");
                    break;

                case TextStyle.Underline:
                    styleContent.Append("<style:text-properties style:text-underline-style=\"solid\" style:text-underline-width=\"auto\" style:text-underline-color=\"font-color\"/>");
                    break;

                case TextStyle.Bold | TextStyle.Italic:
                    styleContent.Append("<style:text-properties fo:font-style=\"italic\" fo:font-weight=\"bold\" style:font-style-asian=\"italic\" style:font-weight-asian=\"bold\" style:font-style-complex=\"italic\" style:font-weight-complex=\"bold\"/>");
                    break;

                case TextStyle.Bold | TextStyle.Underline:
                    styleContent.Append("<style:text-properties style:text-underline-style=\"solid\" style:text-underline-width=\"auto\" style:text-underline-color=\"font-color\" fo:font-weight=\"bold\" style:font-weight-asian=\"bold\" style:font-weight-complex=\"bold\"/>");
                    break;

                case TextStyle.Italic | TextStyle.Underline:
                    styleContent.Append("<style:text-properties fo:font-style=\"italic\" style:text-underline-style=\"solid\" style:text-underline-width=\"auto\" style:text-underline-color=\"font-color\" style:font-style-asian=\"italic\" style:font-style-complex=\"italic\"/>");
                    break;

                case TextStyle.Bold | TextStyle.Italic | TextStyle.Underline:
                    styleContent.Append("<style:text-properties fo:font-style=\"italic\" style:text-underline-style=\"solid\" style:text-underline-width=\"auto\" style:text-underline-color=\"font-color\" fo:font-weight=\"bold\" style:font-style-asian=\"italic\" style:font-weight-asian=\"bold\" style:font-style-complex=\"italic\" style:font-weight-complex=\"bold\"/>");
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(style), style, "Text style not supported");
            }
        }

        /// <summary>
        /// Append a XML end element for style informations
        /// </summary>
        /// <param name="styleContent">The content container for the style informations</param>
        internal static void AppendXmlStyleEnd(in StringBuilder styleContent)
            => styleContent.Append("</style:style>");

        /// <summary>
        /// Return the outline level and style name of a given text style
        /// </summary>
        /// <param name="style">The text style for the style name</param>
        /// <returns>The outline level and style name and </returns>
        /// <exception cref="ArgumentOutOfRangeException">Text style not supported</exception>
        internal static (byte outlineLevel, string styleName) GetStyleAndOutline<TStyle>(in TStyle style)
            where TStyle : notnull, Enum
            => style switch
            {
                TextStyle.Normal                                        => (00, "P1"),
                TextStyle.Bold                                          => (00, "P2"),
                TextStyle.Italic                                        => (00, "P3"),
                TextStyle.Bold | TextStyle.Italic                       => (00, "P4"),
                TextStyle.Underline                                     => (00, "P5"),
                TextStyle.Bold | TextStyle.Underline                    => (00, "P6"),
                TextStyle.Italic | TextStyle.Underline                  => (00, "P7"),
                TextStyle.Bold | TextStyle.Italic | TextStyle.Underline => (00, "P8"),

                TextAlignment.Left                                      => (00, "P21"),
                TextAlignment.Center                                    => (00, "P22"),
                TextAlignment.Right                                     => (00, "P23"),
                TextAlignment.Justify                                   => (00, "P24"),

                HeaderStyle.Title                                       => (00, "Title"),
                HeaderStyle.Subtitle                                    => (00, "Subtitle"),
                HeaderStyle.Signature                                   => (00, "Signature"),
                HeaderStyle.Quotations                                  => (00, "Quotations"),
                HeaderStyle.Endnote                                     => (00, "Endnote"),
                HeaderStyle.Footnote                                    => (00, "Footnote"),

                HeaderStyle.HeadingLevel01                              => (01, "Heading_20_1"),
                HeaderStyle.HeadingLevel02                              => (02, "Heading_20_2"),
                HeaderStyle.HeadingLevel03                              => (03, "Heading_20_3"),
                HeaderStyle.HeadingLevel04                              => (04, "Heading_20_4"),
                HeaderStyle.HeadingLevel05                              => (05, "Heading_20_5"),
                HeaderStyle.HeadingLevel06                              => (06, "Heading_20_6"),
                HeaderStyle.HeadingLevel07                              => (07, "Heading_20_7"),
                HeaderStyle.HeadingLevel08                              => (08, "Heading_20_8"),
                HeaderStyle.HeadingLevel09                              => (09, "Heading_20_9"),
                HeaderStyle.HeadingLevel10                              => (10, "Heading_20_10"),

                _ => throw new ArgumentOutOfRangeException(nameof(style), style, "Text style not supported")
            };

        /// <summary>
        /// Return the name representation for a style for the given table cell
        /// </summary>
        /// <param name="rowNumber">The row number of the current cell</param>
        /// <param name="columnNumber">The column number of the current cell</param>
        /// <param name="columnCount">The column count of the current row</param>
        /// <returns>The name representation of a table column</returns>
        internal static string GetTableCellStyleName(in int rowNumber, in int columnNumber, in int columnCount)
        {
            var number = rowNumber == 1 ? "1" : "2";
            var prefix = columnCount == columnNumber ? "E" : "A";

            return prefix + number;
        }
    }
}
