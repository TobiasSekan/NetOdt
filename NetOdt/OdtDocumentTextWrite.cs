using NetCoreOdt.Enumerations;
using NetCoreOdt.Helper;
using NetOdt.Enumerations;
using System;
using System.Text;

// TODO: consolidate more of the functions of this partial class, use more generic

namespace NetCoreOdt
{
    /// <summary>
    /// Class to create and write ODT documents
    /// </summary>
    public sealed partial class OdtDocument
    {
        /// <summary>
        /// Append a single line with a unformatted value to the document
        /// </summary>
        /// <param name="value">The value to write into the document</param>
        public void Append(in ValueType value)
            => Append(value, TextStyle.Normal);

        /// <summary>
        /// Append a single line with a unformatted text to the document
        /// </summary>
        /// <param name="text">The text to write into the document</param>
        public void Append(in string text)
            => Append(text, TextStyle.Normal);

        /// <summary>
        /// Append the content of the given <see cref="StringBuilder"/> as unformatted text into the document
        /// </summary>
        /// <param name="content">The <see cref="StringBuilder"/> that contains the content for the document</param>
        public void Append(in StringBuilder content)
            => Append(content, TextStyle.Normal);

        /// <summary>
        /// Append a single line with a styled value to the document
        /// </summary>
        /// <param name="value">The value to write into the document</param>
        /// <param name="style">The text style of the value</param>
        public void Append<TStyle>(in ValueType value, in TStyle style) where TStyle : notnull, Enum
        {
            TextContent.Append(GetXmlTextStart(style));
            TextContent.Append(value);
            TextContent.Append(GetXmlTextEnd(style));
        }

        /// <summary>
        /// Append a single line with a styled text to the document
        /// </summary>
        /// <param name="text">The text to write into the document</param>
        /// <param name="style">The text style of the text</param>
        public void Append<TStyle>(in string text, in TStyle style) where TStyle : notnull, Enum
        {
            if(text.Length == 0 || !text.Contains("\n"))
            {
                TextContent.Append(GetXmlTextStart(style));
                TextContent.Append(text);
                TextContent.Append(GetXmlTextEnd(style));
                return;
            }

            foreach(var textBlock in text.Split('\n'))
            {
                TextContent.Append(GetXmlTextStart(style));
                TextContent.Append(textBlock);
                TextContent.Append(GetXmlTextEnd(style));
            }
        }

        /// <summary>
        /// Append the content of the given <see cref="StringBuilder"/> as styled text the document
        /// </summary>
        /// <param name="content">The <see cref="StringBuilder"/> that contains the content for the document</param>
        /// <param name="style">The text style of the content</param>
        public void Append<TStyle>(in StringBuilder content, in TStyle style) where TStyle : notnull, Enum
        {
            if(content.Length == 0 || !StringBuilderHelper.ContainsLineBreaks(content))
            {
                TextContent.Append(GetXmlTextStart(style));
                TextContent.Append(content);
                TextContent.Append(GetXmlTextEnd(style));
                return;
            }

            foreach(var contentBlock in content.ToString().Split('\n'))
            {
                TextContent.Append(GetXmlTextStart(style));
                TextContent.Append(contentBlock);
                TextContent.Append(GetXmlTextEnd(style));
            }
        }

        /// <summary>
        /// Append the given count of empty lines
        /// </summary>
        /// <param name="countOfEmptyLines">The count of empty lines to write</param>
        public void AppendEmptyLines(int countOfEmptyLines)
        {
            for(var count = 0; count < countOfEmptyLines; count++)
            {
                Append(string.Empty, TextStyle.Normal);
            }
        }

        /// <summary>
        /// Return a XML start element for a text passage
        /// </summary>
        /// <typeparam name="TStyle">The <see cref="Type"/> of the style for the text passage</typeparam>
        /// <param name="style">The style for the text passage</param>
        /// <returns>A XML start element for a text passage</returns>
        internal string GetXmlTextStart<TStyle>(in TStyle style) where TStyle : notnull, Enum
        {
            var styleName = StyleHelper.GetStyleName(style);

            return style switch
            {
                HeaderStyle.HeadingLevel01 => $"<text:h text:style-name=\"{styleName}\" text:outline-level=\"1\">",
                HeaderStyle.HeadingLevel02 => $"<text:h text:style-name=\"{styleName}\" text:outline-level=\"2\">",
                HeaderStyle.HeadingLevel03 => $"<text:h text:style-name=\"{styleName}\" text:outline-level=\"3\">",
                HeaderStyle.HeadingLevel04 => $"<text:h text:style-name=\"{styleName}\" text:outline-level=\"4\">",
                HeaderStyle.HeadingLevel05 => $"<text:h text:style-name=\"{styleName}\" text:outline-level=\"5\">",
                HeaderStyle.HeadingLevel06 => $"<text:h text:style-name=\"{styleName}\" text:outline-level=\"6\">",
                HeaderStyle.HeadingLevel07 => $"<text:h text:style-name=\"{styleName}\" text:outline-level=\"7\">",
                HeaderStyle.HeadingLevel08 => $"<text:h text:style-name=\"{styleName}\" text:outline-level=\"8\">",
                HeaderStyle.HeadingLevel09 => $"<text:h text:style-name=\"{styleName}\" text:outline-level=\"9\">",
                HeaderStyle.HeadingLevel10 => $"<text:h text:style-name=\"{styleName}\" text:outline-level=\"10\">",

               _                           => $"<text:p text:style-name=\"{styleName}\">",
            };
        }

        /// <summary>
        /// Return a XML end element for a text passage
        /// </summary>
        /// <returns></returns>
        internal string GetXmlTextEnd<TStyle>(in TStyle style) where TStyle : notnull, Enum
            => style switch
            {
                HeaderStyle.HeadingLevel01 => "</text:h>",
                HeaderStyle.HeadingLevel02 => "</text:h>",
                HeaderStyle.HeadingLevel03 => "</text:h>",
                HeaderStyle.HeadingLevel04 => "</text:h>",
                HeaderStyle.HeadingLevel05 => "</text:h>",
                HeaderStyle.HeadingLevel06 => "</text:h>",
                HeaderStyle.HeadingLevel07 => "</text:h>",
                HeaderStyle.HeadingLevel08 => "</text:h>",
                HeaderStyle.HeadingLevel09 => "</text:h>",
                HeaderStyle.HeadingLevel10 => "</text:h>",

                _                          => "</text:p>"
            };
    }
}
