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
            // P1 - Normal - 000
            styleContent.Append("<style:style style:name=\"P1\" style:family=\"paragraph\" style:parent-style-name=\"Standard\"><style:text-properties/></style:style>");

            // P2 - Bold - 001
            styleContent.Append("<style:style style:name=\"P2\" style:family=\"paragraph\" style:parent-style-name=\"Standard\"><style:text-properties fo:font-weight=\"bold\" style:font-weight-asian=\"bold\" style:font-weight-complex=\"bold\"/></style:style>");

            // P3 - Italic -010
            styleContent.Append("<style:style style:name=\"P3\" style:family=\"paragraph\" style:parent-style-name=\"Standard\"><style:text-properties fo:font-style=\"italic\" style:font-style-asian=\"italic\" style:font-style-complex=\"italic\"/></style:style>");

            // P4 - Bold + Italic - 011
            styleContent.Append("<style:style style:name=\"P4\" style:family=\"paragraph\" style:parent-style-name=\"Standard\"><style:text-properties fo:font-style=\"italic\" fo:font-weight=\"bold\" style:font-style-asian=\"italic\" style:font-weight-asian=\"bold\" style:font-style-complex=\"italic\" style:font-weight-complex=\"bold\"/></style:style>");

            // P5 0b_100 - Underline
            styleContent.Append("<style:style style:name=\"P5\" style:family=\"paragraph\" style:parent-style-name=\"Standard\"><style:text-properties style:text-underline-style=\"solid\" style:text-underline-width=\"auto\" style:text-underline-color=\"font-color\"/></style:style>");

            // P6 - 0b_101 - Bold + Underline
            styleContent.Append("<style:style style:name=\"P6\" style:family=\"paragraph\" style:parent-style-name=\"Standard\"><style:text-properties style:text-underline-style=\"solid\" style:text-underline-width=\"auto\" style:text-underline-color=\"font-color\" fo:font-weight=\"bold\" style:font-weight-asian=\"bold\" style:font-weight-complex=\"bold\"/></style:style>");

            // P7 - 0b_110 - Italic + Underline
            styleContent.Append("<style:style style:name=\"P7\" style:family=\"paragraph\" style:parent-style-name=\"Standard\"><style:text-properties fo:font-style=\"italic\" style:text-underline-style=\"solid\" style:text-underline-width=\"auto\" style:text-underline-color=\"font-color\" style:font-style-asian=\"italic\" style:font-style-complex=\"italic\"/></style:style>");

            // P8 - 0b_111 - Bold + Italic + Underline
            styleContent.Append("<style:style style:name=\"P8\" style:family=\"paragraph\" style:parent-style-name=\"Standard\"><style:text-properties fo:font-style=\"italic\" style:text-underline-style=\"solid\" style:text-underline-width=\"auto\" style:text-underline-color=\"font-color\" fo:font-weight=\"bold\" style:font-style-asian=\"italic\" style:font-weight-asian=\"bold\" style:font-style-complex=\"italic\" style:font-weight-complex=\"bold\"/></style:style>");
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
            styleContent.Append("<style:style style:name=\"Tabelle1.E1\" style:family=\"table-cell\"><style:table-cell-properties fo:padding=\"0.097cm\" fo:border=\"0.05pt solid #000000\"/>");
            styleContent.Append("</style:style>");
            styleContent.Append("<style:style style:name=\"Tabelle1.A2\" style:family=\"table-cell\"><style:table-cell-properties fo:padding=\"0.097cm\" fo:border-left=\"0.05pt solid #000000\" fo:border-right=\"none\" fo:border-top=\"none\" fo:border-bottom=\"0.05pt solid #000000\"/>");
            styleContent.Append("</style:style>");
            styleContent.Append("<style:style style:name=\"Tabelle1.E2\" style:family=\"table-cell\"><style:table-cell-properties fo:padding=\"0.097cm\" fo:border-left=\"0.05pt solid #000000\" fo:border-right=\"0.05pt solid #000000\" fo:border-top=\"none\" fo:border-bottom=\"0.05pt solid #000000\"/>");
            styleContent.Append("</style:style>");
            styleContent.Append("<style:style style:name=\"P1\" style:family=\"paragraph\" style:parent-style-name=\"Table_20_Contents\">");
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
        /// Return the name representation of a given <see cref="TextStyle"/>
        /// </summary>
        /// <param name="style">The text style for the style name</param>
        /// <returns>The name representation of the <see cref="TextStyle"/> </returns>
        internal static string GetStyleName(in TextStyle style)
            => style switch
            {
                TextStyle.Normal                                        => "P1",
                TextStyle.Bold                                          => "P2",
                TextStyle.Italic                                        => "P3",
                TextStyle.Bold | TextStyle.Italic                       => "P4",
                TextStyle.Underline                                     => "P5",
                TextStyle.Bold | TextStyle.Underline                    => "P6",
                TextStyle.Italic | TextStyle.Underline                  => "P7",
                TextStyle.Bold | TextStyle.Italic | TextStyle.Underline => "P8",

                _ => throw new NotSupportedException("Text style not supported")
            };

        /// <summary>
        /// Return the name representation of a given <see cref="HeaderStyle"/>
        /// </summary>
        /// <param name="style">The header style for the style name</param>
        /// <returns>The name representation of the <see cref="HeaderStyle"/> </returns>
        internal static string GetStyleName(in HeaderStyle style)
            => style switch
            {
                HeaderStyle.Title      => "Title",
                HeaderStyle.HeadingLevel01   => "Heading_20_1",
                HeaderStyle.HeadingLevel02   => "Heading_20_2",
                HeaderStyle.HeadingLevel03   => "Heading_20_3",
                HeaderStyle.HeadingLevel04   => "Heading_20_4",
                HeaderStyle.HeadingLevel05   => "Heading_20_5",
                HeaderStyle.HeadingLevel06   => "Heading_20_6",
                HeaderStyle.HeadingLevel07   => "Heading_20_7",
                HeaderStyle.HeadingLevel08   => "Heading_20_8",
                HeaderStyle.HeadingLevel09   => "Heading_20_9",
                HeaderStyle.HeadingLevel10  => "Heading_20_10",
                HeaderStyle.Subtitle   => "Subtitle",
                HeaderStyle.Signature  => "Signature",
                HeaderStyle.Quotations => "Quotations",
                HeaderStyle.Endnote    => "Endnote",
                HeaderStyle.Footnote   => "Footnote",

                _ => throw new NotSupportedException("Header style not supported")
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
