using NetOdt.Enumerations;
using System;
using System.Collections.Generic;

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
        /// <param name="document">The document for the XML style informations</param>
        /// <param name="neededStyles">List that contains all needed styles</param>
        internal static void AddStandardTextStyles(in OdtDocument document, IDictionary<TextStyle, string> neededStyles)
        {
            foreach(var style in neededStyles)
            {
                AppendXmlStyleStart(document, style.Value, StyleFamily.Paragraph, style.Key);
                AppendXmlStyle(document, style.Key);
                AppendXmlStyleEnd(document);
            }
        }

        /// <summary>
        /// Add all needed styles for simple tables
        /// </summary>
        /// <param name="document">The document for the XML style informations</param>
        internal static void AddTableStyles(in OdtDocument document)
        {
            AppendXmlStyleStart(document, "Tabelle1", StyleFamily.Table, TextStyle.None);
            document.StyleContent.Append("<style:table-properties style:width=\"17cm\" table:align=\"margins\"/>");
            AppendXmlStyleEnd(document);

            AppendXmlStyleStart(document, "Tabelle1.A", StyleFamily.TableColumn, TextStyle.None);
            document.StyleContent.Append("<style:table-column-properties style:column-width=\"3.401cm\" style:rel-column-width=\"13107*\"/>");
            AppendXmlStyleEnd(document);

            AppendXmlStyleStart(document, "Tabelle1.A1", StyleFamily.TableCell, TextStyle.None);
            document.StyleContent.Append("<style:table-cell-properties fo:padding=\"0.097cm\" fo:border-left=\"0.05pt solid #000000\" fo:border-right=\"none\" fo:border-top=\"0.05pt solid #000000\" fo:border-bottom=\"0.05pt solid #000000\"/>");
            AppendXmlStyleEnd(document);

            AppendXmlStyleStart(document, "Tabelle1.E1", StyleFamily.TableCell, TextStyle.None);
            document.StyleContent.Append("<style:table-cell-properties fo:padding=\"0.097cm\" fo:border=\"0.05pt solid #000000\"/>");
            AppendXmlStyleEnd(document);

            AppendXmlStyleStart(document, "Tabelle1.A2", StyleFamily.TableCell, TextStyle.None);
            document.StyleContent.Append("<style:table-cell-properties fo:padding=\"0.097cm\" fo:border-left=\"0.05pt solid #000000\" fo:border-right=\"none\" fo:border-top=\"none\" fo:border-bottom=\"0.05pt solid #000000\"/>");
            AppendXmlStyleEnd(document);

            AppendXmlStyleStart(document, "Tabelle1.E2", StyleFamily.TableCell, TextStyle.None);
            document.StyleContent.Append("<style:table-cell-properties fo:padding=\"0.097cm\" fo:border-left=\"0.05pt solid #000000\" fo:border-right=\"0.05pt solid #000000\" fo:border-top=\"none\" fo:border-bottom=\"0.05pt solid #000000\"/>");
            AppendXmlStyleEnd(document);
        }

        /// <summary>
        /// Add all needed styles for simple pictures
        /// </summary>
        /// <param name="document">The document for the XML style informations</param>
        internal static void AddPictureStyle(in OdtDocument document)
        {
            AppendXmlStyleStart(document, "fr1", StyleFamily.Graphic, TextStyle.None);
            document.StyleContent.Append("<style:graphic-properties style:mirror=\"none\" fo:clip=\"rect(0cm, 0cm, 0cm, 0cm)\" draw:luminance=\"0%\" draw:contrast=\"0%\" draw:red=\"0%\" draw:green=\"0%\" draw:blue=\"0%\" draw:gamma=\"100%\" draw:color-inversion=\"false\" draw:image-opacity=\"100%\" draw:color-mode=\"standard\"/>");
            AppendXmlStyleEnd(document);
        }

        /// <summary>
        /// Append a XML start element for XML style informations
        /// </summary>
        /// <param name="document">The document for the XML style informations</param>
        /// <param name="styleName">The name for the style</param>
        /// <param name="styleFamily">The family of the style</param>
        /// <param name="textStyle">The type of the style</param>
        /// <exception cref="ArgumentOutOfRangeException">Style family not supported</exception>
        internal static void AppendXmlStyleStart(in OdtDocument document, in string styleName, in StyleFamily styleFamily, in TextStyle textStyle)
        {
            var styleFamilyName = styleFamily switch
            {
                StyleFamily.Paragraph   => "paragraph",
                StyleFamily.Table       => "table",
                StyleFamily.TableColumn => "table-column",
                StyleFamily.TableCell   => "table-cell",
                StyleFamily.Graphic     => "graphic",
                _                       => throw new ArgumentOutOfRangeException(nameof(styleFamily), styleFamily, "Style family not supported"),
            };

            document.StyleContent.Append($"<");
            document.StyleContent.Append($"style:style");
            document.StyleContent.Append($" style:name=\"{styleName}\"");
            document.StyleContent.Append($" style:family=\"{styleFamilyName}\"");

            switch(styleFamily)
            {
                case StyleFamily.Graphic:
                    document.StyleContent.Append($" style:parent-style-name=\"Graphics\"");
                    break;

                case StyleFamily.Paragraph:
                    document.StyleContent.Append($" style:parent-style-name=\"{GetParentStyleName(textStyle)}\"");
                    break;
            }

            document.StyleContent.Append($">");
        }

        /// <summary>
        /// Append the given text style to the XML style informations
        /// </summary>
        /// <param name="document">The document for the XML style informations</param>
        /// <param name="style">The style for the style informations</param>
        internal static void AppendXmlStyle(in OdtDocument document, TextStyle style)
        {
            AppendTextProperties(document, style);
            AppendParagraphProperties(document, style);
        }

        /// <summary>
        /// Append the given text properties to the XML style informations
        /// </summary>
        /// <param name="document">The document for the XML style informations</param>
        /// <param name="textStyle">The style for the style informations</param>
        internal static void AppendTextProperties(in OdtDocument document, TextStyle textStyle)
        {
            // Note: Don't forget the spaces between the XML properties

            document.StyleContent.Append("<");
            document.StyleContent.Append("style:text-properties");

            document.StyleContent.Append($" style:font-name=\"{document.GlobalFontName}\"");

            if(!HaveParentFontSize(textStyle))
            {
                document.StyleContent.Append($" style:font-size=\"{document.GlobalFontSize}\"");
            }

            if(textStyle.HasFlag(TextStyle.Bold))
            {
                document.StyleContent.Append(" fo:font-weight=\"bold\"");
                document.StyleContent.Append(" style:font-weight-asian=\"bold\"");
                document.StyleContent.Append(" style:font-weight-complex=\"bold\"");
            }
            else
            {
                document.StyleContent.Append(" fo:font-weight=\"normal\"");
                document.StyleContent.Append(" style:font-weight-asian=\"normal\"");
                document.StyleContent.Append(" style:font-weight-complex=\"normal\"");
            }

            if(textStyle.HasFlag(TextStyle.Italic))
            {
                document.StyleContent.Append(" fo:font-style=\"italic\"");
                document.StyleContent.Append(" style:font-style-asian=\"italic\"");
                document.StyleContent.Append(" style:font-style-complex=\"italic\"");
            }
            else
            {
                document.StyleContent.Append(" fo:font-style=\"normal\"");
                document.StyleContent.Append(" style:font-style-asian=\"normal\"");
                document.StyleContent.Append(" style:font-style-complex=\"normal\"");
            }

            if(textStyle.HasFlag(TextStyle.UnderlineSingle))
            {
                document.StyleContent.Append(" style:text-underline-style=\"solid\"");
                document.StyleContent.Append(" style:text-underline-width=\"auto\"");
                document.StyleContent.Append(" style:text-underline-color=\"font-color\"");
            }
            else if(textStyle.HasFlag(TextStyle.UnderlineDouble))
            {
                document.StyleContent.Append(" style:text-underline-style=\"double\"");
                document.StyleContent.Append(" style:text-underline-width=\"auto\"");
                document.StyleContent.Append(" style:text-underline-color=\"font-color\"");
            }
            else if(textStyle.HasFlag(TextStyle.UnderlineBold))
            {
                document.StyleContent.Append(" style:text-underline-style=\"solid\"");
                document.StyleContent.Append(" style:text-underline-width=\"bold\"");
                document.StyleContent.Append(" style:text-underline-color=\"font-color\"");
            }
            else if(textStyle.HasFlag(TextStyle.UnderlineDotted))
            {
                document.StyleContent.Append(" style:text-underline-style=\"dotted\"");
                document.StyleContent.Append(" style:text-underline-width=\"auto\"");
                document.StyleContent.Append(" style:text-underline-color=\"font-color\"");
            }
            else if(textStyle.HasFlag(TextStyle.UnderlineDottedBold))
            {
                document.StyleContent.Append(" style:text-underline-style=\"dotted\"");
                document.StyleContent.Append(" style:text-underline-width=\"bold\"");
                document.StyleContent.Append(" style:text-underline-color=\"font-color\"");
            }
            else if(textStyle.HasFlag(TextStyle.UnderlineDash))
            {
                document.StyleContent.Append(" style:text-underline-style=\"dash\"");
                document.StyleContent.Append(" style:text-underline-width=\"auto\"");
                document.StyleContent.Append(" style:text-underline-color=\"font-color\"");
            }
            else if(textStyle.HasFlag(TextStyle.UnderlineLongDash))
            {
                document.StyleContent.Append(" style:text-underline-style=\"long-dash\"");
                document.StyleContent.Append(" style:text-underline-width=\"auto\"");
                document.StyleContent.Append(" style:text-underline-color=\"font-color\"");
            }
            else if(textStyle.HasFlag(TextStyle.UnderlineDotDash))
            {
                document.StyleContent.Append(" style:text-underline-style=\"dot-dash\"");
                document.StyleContent.Append(" style:text-underline-width=\"auto\"");
                document.StyleContent.Append(" style:text-underline-color=\"font-color\"");
            }
            else if(textStyle.HasFlag(TextStyle.UnderlineDotDotDash))
            {
                document.StyleContent.Append(" style:text-underline-style=\"dot-dot-dash\"");
                document.StyleContent.Append(" style:text-underline-width=\"auto\"");
                document.StyleContent.Append(" style:text-underline-color=\"font-color\"");
            }
            else if(textStyle.HasFlag(TextStyle.UnderlineWave))
            {
                document.StyleContent.Append(" style:text-underline-style=\"wave\"");
                document.StyleContent.Append(" style:text-underline-width=\"auto\"");
                document.StyleContent.Append(" style:text-underline-color=\"font-color\"");
            }
            else
            {
                document.StyleContent.Append(" style:text-underline-style=\"none\"");
            }

            if(textStyle.HasFlag(TextStyle.Stroke))
            {
                document.StyleContent.Append(" style:text-line-through-style=\"solid\"");
                document.StyleContent.Append(" style:text-line-through-type=\"single\"");
            }
            else
            {
                document.StyleContent.Append(" style:text-line-through-style=\"none\"");
                document.StyleContent.Append(" style:text-line-through-type=\"none\"");
            }

            // A text paragraph can only use superscript or subscript not both at same time
            if(textStyle.HasFlag(TextStyle.Subscript))
            {
                document.StyleContent.Append(" style:text-position=\"sub 58%\"");
            }
            else if(textStyle.HasFlag(TextStyle.Superscript))
            {
                document.StyleContent.Append(" style:text-position=\"super 58%\"");
            }

            document.StyleContent.Append("/>");
        }

        /// <summary>
        /// Append the given paragraph properties to the XML style informations
        /// </summary>
        /// <param name="document">The document for the XML style informations</param>
        /// <param name="style">The style for the style informations</param>
        internal static void AppendParagraphProperties(in OdtDocument document, TextStyle style)
        {
            if(!style.HasFlag(TextStyle.Left)
            && !style.HasFlag(TextStyle.Center)
            && !style.HasFlag(TextStyle.Right)
            && !style.HasFlag(TextStyle.Justify)
            && !style.HasFlag(TextStyle.PageBreak))
            {
                return;
            }

            // Note: Don't forget the spaces between the XML properties

            document.StyleContent.Append("<");
            document.StyleContent.Append("style:paragraph-properties");

            if(style.HasFlag(TextStyle.Left))
            {
                document.StyleContent.Append(" fo:text-align=\"start\"");
                document.StyleContent.Append(" style:justify-single-word=\"false\"");
            }

            if(style.HasFlag(TextStyle.Center))
            {
                document.StyleContent.Append(" fo:text-align=\"center\"");
                document.StyleContent.Append(" style:justify-single-word=\"false\"");
            }

            if(style.HasFlag(TextStyle.Right))
            {
                document.StyleContent.Append(" fo:text-align=\"end\"");
                document.StyleContent.Append(" style:justify-single-word=\"false\"");
            }

            if(style.HasFlag(TextStyle.Justify))
            {
                document.StyleContent.Append(" fo:text-align=\"justify\"");
                document.StyleContent.Append(" style:justify-single-word=\"false\"");
            }

            if(style.HasFlag(TextStyle.Justify))
            {
                document.StyleContent.Append(" fo:break-before=\"page\"");
            }

            document.StyleContent.Append("/>");
        }

        /// <summary>
        /// Append a XML end element for XML style informations
        /// </summary>
        /// <param name="document">The document for the XML style informations</param>
        internal static void AppendXmlStyleEnd(in OdtDocument document)
            => document.StyleContent.Append("</style:style>");

        /// <summary>
        /// Return the outline level of a given text style
        /// </summary>
        /// <param name="textStyle">The text style for the outline level</param>
        /// <returns>A outline level</returns>
        internal static byte GetOutlineLevel(in TextStyle textStyle)
            => textStyle switch
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

        /// <summary>
        /// Return the parent style name of a given text style
        /// </summary>
        /// <param name="textStyle">The text style for the parent style name</param>
        /// <returns>A parent style name</returns>
        internal static string GetParentStyleName(TextStyle textStyle)
            => textStyle switch
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

        /// <summary>
        /// Check if the given text style have a parent font size
        /// </summary>
        /// <param name="textStyle">The text style to check</param>
        /// <returns><see langword="true"/> if the given text style have a parent font size, otherwise <see langword="false"/></returns>
        internal static bool HaveParentFontSize(TextStyle textStyle)
            => textStyle switch
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
    }
}
