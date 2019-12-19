using NetCoreOdt.Enumerations;
using NetCoreOdt.Helper;
using System;
using System.Text;

namespace NetCoreOdt
{
    /// <summary>
    /// Class to create and write ODT documents
    /// </summary>
    public sealed partial class OdtDocument
    {
        /// <summary>
        /// Write a single line with a unformatted value to the document
        /// </summary>
        /// <param name="value">The value to write into the document</param>
        public void Write(in ValueType value)
            => Write(value, TextStyle.Normal);

        /// <summary>
        /// Write a single line with a unformatted text to the document (Note: line breaks "\n" will currently not working)
        /// </summary>
        /// <param name="text">The text to write into the document</param>
        public void Write(in string text)
            => Write(text, TextStyle.Normal);

        /// <summary>
        /// Write the content of the given <see cref="StringBuilder"/> as unformatted text into the document (Note: line breaks "\n" will currently not working)
        /// </summary>
        /// <param name="content">The <see cref="StringBuilder"/> that contains the content for the document</param>
        public void Write(in StringBuilder content)
            => Write(content, TextStyle.Normal);

        /// <summary>
        /// Write a single line with a styled value to the document
        /// </summary>
        /// <param name="value">The value to write into the document</param>
        /// <param name="style">The text style of the value</param>
        public void Write(in ValueType value, in TextStyle style)
        {
            TextContent.Append($"<text:p text:style-name=\"{StyleHelper.GetStyleName(style)}\">");
            TextContent.Append(value);
            TextContent.Append("</text:p>");
        }

        /// <summary>
        /// Write a single line with a styled text to the document (Note: line breaks "\n" will currently not working)
        /// </summary>
        /// <param name="text">The text to write into the document</param>
        /// <param name="style">The text style of the text</param>
        public void Write(in string text, in TextStyle style)
        {
            var styleName = StyleHelper.GetStyleName(style);

            if(!text.Contains("\n"))
            {
                TextContent.Append($"<text:p text:style-name=\"{styleName}\">");
                TextContent.Append(text);
                TextContent.Append("</text:p>");
                return;
            }

            foreach(var textBlock in text.Split('\n'))
            {
                TextContent.Append($"<text:p text:style-name=\"{styleName}\">");
                TextContent.Append(textBlock);
                TextContent.Append("</text:p>");
            }
        }

        /// <summary>
        /// Write the content of the given <see cref="StringBuilder"/> as styled text the document (Note: line breaks "\n" will currently not working)
        /// </summary>
        /// <param name="content">The <see cref="StringBuilder"/> that contains the content for the document</param>
        /// <param name="style">The text style of the content</param>
        public void Write(in StringBuilder content, in TextStyle style)
        {
            var styleName          = StyleHelper.GetStyleName(style);
            var containsLineBreaks = false;

            for(var index = 0; index < content.Length; index++)
            {
                if(content[index] == '\n')
                {
                    containsLineBreaks = true;
                    break;
                }
            }

            if(!containsLineBreaks)
            {
                TextContent.Append($"<text:p text:style-name=\"{styleName}\">");
                TextContent.Append(content);
                TextContent.Append("</text:p>");
                return;
            }

            foreach(var contentBlock in content.ToString().Split('\n'))
            {
                TextContent.Append($"<text:p text:style-name=\"{styleName}\">");
                TextContent.Append(contentBlock);
                TextContent.Append("</text:p>");
            }
        }
    }
}
