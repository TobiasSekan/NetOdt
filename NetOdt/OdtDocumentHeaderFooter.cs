using NetOdt.Enumerations;
using NetOdt.Helper;
using System;
using System.Text;

namespace NetOdt
{
    /// <summary>
    /// Class to set the global header and footer
    /// </summary>
    public sealed partial class OdtDocument
    {
        /// <summary>
        /// The raw content of the master style
        /// </summary>
        internal StringBuilder MasterStyle { get; }

        /// <summary>
        /// The raw content of the header
        /// </summary>
        internal StringBuilder HeaderContent { get; }

        /// <summary>
        /// The raw content of the footer
        /// </summary>
        internal StringBuilder FooterContent { get; }

        /// <summary>
        /// Set the given content as header
        /// </summary>
        /// <typeparam name="TValue">The type of the header content</typeparam>
        /// <param name="value">The content for the header</param>
        /// <exception cref="NotSupportedException">Style-sheets inside the header not supported</exception>
        public void SetHeader<TValue>(in TValue value)
            where TValue : notnull
            => SetHeader(value, TextStyle.None);

        /// <summary>
        /// Set the given content with the given style as header
        /// </summary>
        /// <typeparam name="TValue">The type of the header content</typeparam>
        /// <param name="value">The content for the header</param>
        /// <param name="textStyle">The text style for the header</param>
        /// <exception cref="NotSupportedException">Style-sheets inside the header not supported</exception>
        public void SetHeader<TValue>(in TValue value, TextStyle textStyle)
            where TValue : notnull
        {
            switch(textStyle)
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
                case TextStyle.Title:
                case TextStyle.Subtitle:
                case TextStyle.Signature:
                case TextStyle.Quotations:
                case TextStyle.Endnote:
                case TextStyle.Footnote:
                    throw new NotSupportedException("Style-sheets inside the header not supported");
            }

            MasterStyleCount++;

            StyleHelper.AppendXmlStyleStart(MasterStyle, $"MP{MasterStyleCount}", StyleFamily.Header, textStyle);
            StyleHelper.AppendXmlStyle(MasterStyle, textStyle, GlobalFontName, GlobalFontSize);
            StyleHelper.AppendXmlStyleEnd(MasterStyle);

            // TODO: don't use hard coded name for style informations

            HeaderContent.Append($"<text:p text:style-name=\"MP{MasterStyleCount}\">");
            HeaderContent.Append(value);
            HeaderContent.Append("</text:p>");
        }

        /// <summary>
        /// Set the given content as footer
        /// </summary>
        /// <typeparam name="TValue">The type of the footer content</typeparam>
        /// <param name="value">The content for the footer</param>
        /// <exception cref="NotSupportedException">Style-sheets inside the footer not supported</exception>
        public void SetFooter<TValue>(in TValue value)
            where TValue : notnull
            => SetFooter(value, TextStyle.None);

        /// <summary>
        /// Set the given content with the given style as footer
        /// </summary>
        /// <typeparam name="TValue">The type of the footer content</typeparam>
        /// <param name="value">The content for the footer</param>
        /// <param name="textStyle">The text style for the footer</param>
        /// <exception cref="NotSupportedException">Style-sheets inside the footer not supported</exception>
        public void SetFooter<TValue>(in TValue value, TextStyle textStyle)
            where TValue : notnull
        {
            switch(textStyle)
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
                case TextStyle.Title:
                case TextStyle.Subtitle:
                case TextStyle.Signature:
                case TextStyle.Quotations:
                case TextStyle.Endnote:
                case TextStyle.Footnote:
                    throw new NotSupportedException("Style-sheets inside the footer not supported");
            }

            MasterStyleCount++;

            StyleHelper.AppendXmlStyleStart(MasterStyle, $"MP{MasterStyleCount}", StyleFamily.Footer, textStyle);
            StyleHelper.AppendXmlStyle(MasterStyle, textStyle, GlobalFontName, GlobalFontSize);
            StyleHelper.AppendXmlStyleEnd(MasterStyle);

            // TODO: don't use hard coded name for style informations

            FooterContent.Append($"<text:p text:style-name=\"MP{MasterStyleCount}\">");
            FooterContent.Append(value);
            FooterContent.Append("</text:p>");
        }
    }
}
