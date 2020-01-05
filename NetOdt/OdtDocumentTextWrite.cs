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
        public void Append<TValue>(in TValue value)
            where TValue : notnull
            => Append(value, TextStyle.Normal);

        /// <summary>
        /// Append a single line with a styled value to the document
        /// </summary>
        /// <param name="value">The value to write into the document</param>
        /// <param name="style">The style of the value</param>
        public void Append<TValue, TStyle>(in TValue value, in TStyle style)
            where TValue : notnull
            where TStyle : notnull, Enum
        {
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
        public void AppendEmptyLines(int countOfEmptyLines)
        {
            for(var count = 0; count < countOfEmptyLines; count++)
            {
                Append(string.Empty, TextStyle.Normal);
            }
        }

        /// <summary>
        /// Append a XML start element for a text passage
        /// </summary>
        /// <typeparam name="TStyle">The <see cref="Type"/> of the style for the text passage</typeparam>
        /// <param name="style">The style for the text passage</param>
        internal void AppendXmlTextStart<TStyle>(in TStyle style) where TStyle : notnull, Enum
        {
            var styleData = StyleHelper.GetStyleAndOutline(style);

            switch(style)
            {
                case HeaderStyle.HeadingLevel01:
                case HeaderStyle.HeadingLevel02:
                case HeaderStyle.HeadingLevel03:
                case HeaderStyle.HeadingLevel04:
                case HeaderStyle.HeadingLevel05:
                case HeaderStyle.HeadingLevel06:
                case HeaderStyle.HeadingLevel07:
                case HeaderStyle.HeadingLevel08:
                case HeaderStyle.HeadingLevel09:
                case HeaderStyle.HeadingLevel10:
                    TextContent.Append($"<text:h text:style-name=\"{styleData.Name}\" text:outline-level=\"{styleData.Level}\">");
                    break;

                default:
                    TextContent.Append($"<text:p text:style-name=\"{styleData.Name}\">");
                    break;
            }
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
        internal void AppendXmlTextEnd<TStyle>(in TStyle style) where TStyle : notnull, Enum
        {
            switch(style)
            {
                case HeaderStyle.HeadingLevel01:
                case HeaderStyle.HeadingLevel02:
                case HeaderStyle.HeadingLevel03:
                case HeaderStyle.HeadingLevel04:
                case HeaderStyle.HeadingLevel05:
                case HeaderStyle.HeadingLevel06:
                case HeaderStyle.HeadingLevel07:
                case HeaderStyle.HeadingLevel08:
                case HeaderStyle.HeadingLevel09:
                case HeaderStyle.HeadingLevel10:
                    TextContent.Append("</text:h>");
                    break;

                default:
                    TextContent.Append("</text:p>");
                    break;
            }
        }
    }
}
