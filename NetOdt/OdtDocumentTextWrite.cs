using NetOdt.Class;
using NetOdt.Enumerations;
using NetOdt.Helper;
using System;
using System.Text;

namespace NetOdt
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
        public void AppendLine<TValue>(in TValue value)
            where TValue : notnull
            => AppendLine(value, TextStyle.None);

        /// <summary>
        /// Append a single line with a styled value to the document
        /// </summary>
        /// <param name="value">The value to write into the document</param>
        /// <param name="textStyle">The style of the value</param>
        public void AppendLine<TValue>(in TValue value, in TextStyle textStyle)
            where TValue : notnull
        {
            var style = TryToAddStyle(textStyle);

            if(value is StringBuilder content)
            {
                if(content.Length == 0 || !StringBuilderHelper.ContainsLineBreaks(content))
                {
                    AppendXmlTextStart(style);
                    AppendValue(content);
                    AppendXmlTextEnd(style);
                    return;
                }

                foreach(var contentBlock in content.ToString().Split('\n'))
                {
                    AppendXmlTextStart(style);
                    AppendValue(contentBlock);
                    AppendXmlTextEnd(style);
                }

                return;
            }

            if(value is string text)
            {
                if(text.Length == 0 || !text.Contains("\n"))
                {
                    AppendXmlTextStart(style);
                    AppendValue(text);
                    AppendXmlTextEnd(style);
                    return;
                }

                foreach(var textBlock in text.Split('\n'))
                {
                    AppendXmlTextStart(style);
                    AppendValue(textBlock);
                    AppendXmlTextEnd(style);
                }

                return;
            }

            AppendXmlTextStart(style);
            AppendValue(value);
            AppendXmlTextEnd(style);
        }

        /// <summary>
        /// Append the given count of empty lines
        /// </summary>
        /// <param name="countOfEmptyLines">The count of empty lines to write</param>
        public void AppendEmptyLines(in int countOfEmptyLines)
        {
            for(var count = 0; count < countOfEmptyLines; count++)
            {
                AppendLine(string.Empty, TextStyle.None);
            }
        }

        /// <summary>
        /// Append a XML start element for a text passage
        /// </summary>
        /// <param name="style">The style for the text passage</param>
        internal void AppendXmlTextStart(in Style style)
        {
            TextContent.Append("<");

            switch(style.TextStyle)
            {
                case TextStyle.HeadingLevel01:
                case TextStyle.HeadingLevel02:
                case TextStyle.HeadingLevel03:
                case TextStyle.HeadingLevel04:
                case TextStyle.HeadingLevel05:
                case TextStyle.HeadingLevel06:
                case TextStyle.HeadingLevel07:
                case TextStyle.HeadingLevel08:
                case TextStyle.HeadingLevel09:
                case TextStyle.HeadingLevel10:
                    TextContent.Append("text:h");
                    TextContent.Append($" text:style-name=\"{style.Name}\"");
                    TextContent.Append($" text:outline-level=\"{style.OutlineLevel}\"");
                    break;

                default:
                    TextContent.Append("text:p");
                    TextContent.Append($" text:style-name=\"{style.Name}\"");
                    break;
            }

            TextContent.Append(">");
        }

        /// <summary>
        /// Append the value of a text passage
        /// </summary>
        /// <typeparam name="TValue">The <see cref="Type"/> of the value</typeparam>
        /// <param name="value">The value for the text passage</param>
        internal void AppendValue<TValue>(in TValue value)
            where TValue : notnull
        {
            switch(value)
            {
                case bool boolValue:
                    TextContent.Append(boolValue);
                    break;

                case byte byteValue:
                    TextContent.Append(byteValue);
                    break;

                case sbyte sbyteValue:
                    TextContent.Append(sbyteValue);
                    break;

                case char charValue:
                    TextContent.Append(charValue);
                    break;

                case short shortValue:
                    TextContent.Append(shortValue);
                    break;

                case ushort ushortValue:
                    TextContent.Append(ushortValue);
                    break;

                case int intValue:
                    TextContent.Append(intValue);
                    break;

                case uint uintValue:
                    TextContent.Append(uintValue);
                    break;

                case long longValue:
                    TextContent.Append(longValue);
                    break;

                case ulong ulongValue:
                    TextContent.Append(ulongValue);
                    break;

                case float floatValue:
                    TextContent.Append(floatValue);
                    break;

                case double doubleValue:
                    TextContent.Append(doubleValue);
                    break;

                case decimal decimalValue:
                    TextContent.Append(decimalValue);
                    break;

                case string stringValue:
                    TextContent.Append(stringValue);
                    break;

                // there is no overload that takes a StringBuilder in .NET Standard 2.0
                // but in .Net Standard 2.1 "StringBuilder" have a overload for "StringBuilder"
                //
                case StringBuilder stringBuilder:
                    TextContent.Append(stringBuilder);
                    break;

                // Pattern matching doesn't support "ReadOnlySpan<char>",
                // but in .Net Standard 2.1 "StringBuilder" have a overload for "ReadOnlySpan<char>"
                //
                // case ReadOnlySpan<char> readOnlySpan:
                //     TextContent.Append(readOnlySpan);
                //     break;

                default:
                    TextContent.Append(value);
                    break;
            }
        }

        /// <summary>
        /// Append a XML end element for a text passage
        /// </summary>
        /// <param name="style">The style of this text passage</param>
        internal void AppendXmlTextEnd(in Style style)
        {
            switch(style.TextStyle)
            {
                case TextStyle.HeadingLevel01:
                case TextStyle.HeadingLevel02:
                case TextStyle.HeadingLevel03:
                case TextStyle.HeadingLevel04:
                case TextStyle.HeadingLevel05:
                case TextStyle.HeadingLevel06:
                case TextStyle.HeadingLevel07:
                case TextStyle.HeadingLevel08:
                case TextStyle.HeadingLevel09:
                case TextStyle.HeadingLevel10:
                    TextContent.Append("</text:h>");
                    break;

                default:
                    TextContent.Append("</text:p>");
                    break;
            }
        }
    }
}
