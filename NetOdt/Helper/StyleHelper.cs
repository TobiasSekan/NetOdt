using NetCoreOdt.Enumerations;
using NetOdt.Enumerations;
using System;
using System.Text;

namespace NetCoreOdt.Helper
{
    /// <summary>
    /// Helper class to easier work with ODT document styles
    /// </summary>
    internal static class StyleHelper
    {
        /// <summary>
        /// Add all needed styles for all <see cref="TextStyle"/> combinations to the style content
        /// </summary>
        internal static void AddStandardTextStyles(in StringBuilder styleContent)
        {
            // P1 - 0b_000 - Normal
            styleContent.Append("<style:style style:name=\"P1\" style:family=\"paragraph\" style:parent-style-name=\"Standard\">");
            styleContent.Append("<style:text-properties/>");
            styleContent.Append("</style:style>");

            // P2 - 0b_001 - Bold
            styleContent.Append("<style:style style:name=\"P2\" style:family=\"paragraph\" style:parent-style-name=\"Standard\">");
            styleContent.Append("<style:text-properties fo:font-weight=\"bold\" style:font-weight-asian=\"bold\" style:font-weight-complex=\"bold\"/>");
            styleContent.Append("</style:style>");

            // P3 - 0b_010 - Italic
            styleContent.Append("<style:style style:name=\"P3\" style:family=\"paragraph\" style:parent-style-name=\"Standard\">");
            styleContent.Append("<style:text-properties fo:font-style=\"italic\" style:font-style-asian=\"italic\" style:font-style-complex=\"italic\"/>");
            styleContent.Append("</style:style>");

            // P4 - 0b_011 - Bold + Italic
            styleContent.Append("<style:style style:name=\"P4\" style:family=\"paragraph\" style:parent-style-name=\"Standard\">");
            styleContent.Append("<style:text-properties fo:font-style=\"italic\" fo:font-weight=\"bold\" style:font-style-asian=\"italic\" style:font-weight-asian=\"bold\" style:font-style-complex=\"italic\" style:font-weight-complex=\"bold\"/>");
            styleContent.Append("</style:style>");

            // P5 0b_100 - Underline
            styleContent.Append("<style:style style:name=\"P5\" style:family=\"paragraph\" style:parent-style-name=\"Standard\">");
            styleContent.Append("<style:text-properties style:text-underline-style=\"solid\" style:text-underline-width=\"auto\" style:text-underline-color=\"font-color\"/>");
            styleContent.Append("</style:style>");

            // P6 - 0b_101 - Bold + Underline
            styleContent.Append("<style:style style:name=\"P6\" style:family=\"paragraph\" style:parent-style-name=\"Standard\">");
            styleContent.Append("<style:text-properties style:text-underline-style=\"solid\" style:text-underline-width=\"auto\" style:text-underline-color=\"font-color\" fo:font-weight=\"bold\" style:font-weight-asian=\"bold\" style:font-weight-complex=\"bold\"/>");
            styleContent.Append("</style:style>");

            // P7 - 0b_110 - Italic + Underline
            styleContent.Append("<style:style style:name=\"P7\" style:family=\"paragraph\" style:parent-style-name=\"Standard\">");
            styleContent.Append("<style:text-properties fo:font-style=\"italic\" style:text-underline-style=\"solid\" style:text-underline-width=\"auto\" style:text-underline-color=\"font-color\" style:font-style-asian=\"italic\" style:font-style-complex=\"italic\"/>");
            styleContent.Append("</style:style>");

            // P8 - 0b_111 - Bold + Italic + Underline
            styleContent.Append("<style:style style:name=\"P8\" style:family=\"paragraph\" style:parent-style-name=\"Standard\">");
            styleContent.Append("<style:text-properties fo:font-style=\"italic\" style:text-underline-style=\"solid\" style:text-underline-width=\"auto\" style:text-underline-color=\"font-color\" fo:font-weight=\"bold\" style:font-style-asian=\"italic\" style:font-weight-asian=\"bold\" style:font-style-complex=\"italic\" style:font-weight-complex=\"bold\"/>");
            styleContent.Append("</style:style>");
        }

        /// <summary>
        /// Add all needed styles for all <see cref="TextAlignment"/> combinations to the style content
        /// </summary>
        internal static void AddTextAlignmentStyles(in StringBuilder styleContent)
        {
            // Left text alignment
            styleContent.Append("<style:style style:name=\"P21\" style:family=\"paragraph\" style:parent-style-name=\"Standard\">");
            styleContent.Append("<style:paragraph-properties fo:text-align=\"start\" style:justify-single-word=\"false\"/>");
            styleContent.Append("</style:style>");

            // Centered text
            styleContent.Append("<style:style style:name=\"P22\" style:family=\"paragraph\" style:parent-style-name=\"Standard\">");
            styleContent.Append("<style:paragraph-properties fo:text-align=\"center\" style:justify-single-word=\"false\"/>");
            styleContent.Append("</style:style>");

            // Right text alignment
            styleContent.Append("<style:style style:name=\"P23\" style:family=\"paragraph\" style:parent-style-name=\"Standard\">");
            styleContent.Append("<style:paragraph-properties fo:text-align=\"end\" style:justify-single-word=\"false\"/>");
            styleContent.Append("</style:style>");

            // Full justification
            styleContent.Append("<style:style style:name=\"P24\" style:family=\"paragraph\" style:parent-style-name=\"Standard\">");
            styleContent.Append("<style:paragraph-properties fo:text-align=\"justify\" style:justify-single-word=\"false\"/>");
            styleContent.Append("</style:style>");
        }

        /// <summary>
        /// Add all needed styles for simple tables
        /// </summary>
        internal static void AddTableStyles(in StringBuilder styleContent)
        {
            styleContent.Append("<style:style style:name=\"Tabelle1\" style:family=\"table\">");
            styleContent.Append("<style:table-properties style:width=\"17cm\" table:align=\"margins\"/>");
            styleContent.Append("</style:style>");

            styleContent.Append("<style:style style:name=\"Tabelle1.A\" style:family=\"table-column\">");
            styleContent.Append("<style:table-column-properties style:column-width=\"3.401cm\" style:rel-column-width=\"13107*\"/>");
            styleContent.Append("</style:style>");

            styleContent.Append("<style:style style:name=\"Tabelle1.A1\" style:family=\"table-cell\">");
            styleContent.Append("<style:table-cell-properties fo:padding=\"0.097cm\" fo:border-left=\"0.05pt solid #000000\" fo:border-right=\"none\" fo:border-top=\"0.05pt solid #000000\" fo:border-bottom=\"0.05pt solid #000000\"/>");
            styleContent.Append("</style:style>");

            styleContent.Append("<style:style style:name=\"Tabelle1.E1\" style:family=\"table-cell\">");
            styleContent.Append("<style:table-cell-properties fo:padding=\"0.097cm\" fo:border=\"0.05pt solid #000000\"/>");
            styleContent.Append("</style:style>");

            styleContent.Append("<style:style style:name=\"Tabelle1.A2\" style:family=\"table-cell\">");
            styleContent.Append("<style:table-cell-properties fo:padding=\"0.097cm\" fo:border-left=\"0.05pt solid #000000\" fo:border-right=\"none\" fo:border-top=\"none\" fo:border-bottom=\"0.05pt solid #000000\"/>");
            styleContent.Append("</style:style>");

            styleContent.Append("<style:style style:name=\"Tabelle1.E2\" style:family=\"table-cell\">");
            styleContent.Append("<style:table-cell-properties fo:padding=\"0.097cm\" fo:border-left=\"0.05pt solid #000000\" fo:border-right=\"0.05pt solid #000000\" fo:border-top=\"none\" fo:border-bottom=\"0.05pt solid #000000\"/>");
            styleContent.Append("</style:style>");
        }

        /// <summary>
        /// Add all needed styles for simple pictures
        /// </summary>
        internal static void AddPictureStyle(in StringBuilder styleContent)
        {
            styleContent.Append("<style:style style:name=\"fr1\" style:family=\"graphic\" style:parent-style-name=\"Graphics\">");
            styleContent.Append("<style:graphic-properties style:mirror=\"none\" fo:clip=\"rect(0cm, 0cm, 0cm, 0cm)\" draw:luminance=\"0%\" draw:contrast=\"0%\" draw:red=\"0%\" draw:green=\"0%\" draw:blue=\"0%\" draw:gamma=\"100%\" draw:color-inversion=\"false\" draw:image-opacity=\"100%\" draw:color-mode=\"standard\"/>");
            styleContent.Append("</style:style>");
        }

        /// <summary>
        /// Return the style name and the outline level of a given text style
        /// </summary>
        /// <param name="style">The text style for the style name</param>
        /// <returns>The style name and the outline level</returns>
        internal static (string Name, byte Level) GetStyleAndOutline<TStyle>(in TStyle style)
            where TStyle : notnull, Enum
            => style switch
            {
                TextStyle.Normal                                        => ("P1", 0),
                TextStyle.Bold                                          => ("P2", 0),
                TextStyle.Italic                                        => ("P3", 0),
                TextStyle.Bold | TextStyle.Italic                       => ("P4", 0),
                TextStyle.Underline                                     => ("P5", 0),
                TextStyle.Bold | TextStyle.Underline                    => ("P6", 0),
                TextStyle.Italic | TextStyle.Underline                  => ("P7", 0),
                TextStyle.Bold | TextStyle.Italic | TextStyle.Underline => ("P8", 0),

                TextAlignment.Left                                      => ("P21", 0),
                TextAlignment.Center                                    => ("P22", 0),
                TextAlignment.Right                                     => ("P23", 0),
                TextAlignment.Justify                                   => ("P24", 0),

                HeaderStyle.Title                                       => ("Title", 0),
                HeaderStyle.Subtitle                                    => ("Subtitle", 0),
                HeaderStyle.Signature                                   => ("Signature", 0),
                HeaderStyle.Quotations                                  => ("Quotations", 0),
                HeaderStyle.Endnote                                     => ("Endnote", 0),
                HeaderStyle.Footnote                                    => ("Footnote", 0),

                HeaderStyle.HeadingLevel01                              => ("Heading_20_1", 1),
                HeaderStyle.HeadingLevel02                              => ("Heading_20_2", 2),
                HeaderStyle.HeadingLevel03                              => ("Heading_20_3", 3),
                HeaderStyle.HeadingLevel04                              => ("Heading_20_4", 4),
                HeaderStyle.HeadingLevel05                              => ("Heading_20_5", 5),
                HeaderStyle.HeadingLevel06                              => ("Heading_20_6", 6),
                HeaderStyle.HeadingLevel07                              => ("Heading_20_7", 7),
                HeaderStyle.HeadingLevel08                              => ("Heading_20_8", 8),
                HeaderStyle.HeadingLevel09                              => ("Heading_20_9", 9),
                HeaderStyle.HeadingLevel10                              => ("Heading_20_10", 10),

                _ => throw new NotSupportedException("Text style not supported")
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
