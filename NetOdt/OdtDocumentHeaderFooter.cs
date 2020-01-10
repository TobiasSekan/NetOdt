using NetOdt.Class;
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
        /// Set the given content as header (use stand style and global font)
        /// </summary>
        /// <typeparam name="TValue">The type of the header content</typeparam>
        /// <param name="value">The content for the header</param>
        /// <exception cref="NotSupportedException">Style-sheets inside the header not supported</exception>
        public void SetHeader<TValue>(in TValue value)
            where TValue : notnull
            => SetHeader(value, TextStyle.None, GlobalFontName, GlobalFontSize);

        /// <summary>
        /// Set the given content with the given style as header (use global font)
        /// </summary>
        /// <typeparam name="TValue">The type of the header content</typeparam>
        /// <param name="value">The content for the header</param>
        /// <param name="textStyle">The text style for the header</param>
        /// <exception cref="NotSupportedException">Style-sheets inside the header not supported</exception>
        public void SetHeader<TValue>(in TValue value, TextStyle textStyle)
            where TValue : notnull
            => SetHeader(value, textStyle, GlobalFontName, GlobalFontSize);

        /// <summary>
        /// Set the given content with the given style and font as header
        /// </summary>
        /// <typeparam name="TValue">The type of the header content</typeparam>
        /// <param name="value">The content for the header</param>
        /// <param name="textStyle">The text style for the header</param>
        /// <param name="fontName">The font (name) for the header</param>
        /// <param name="fontSize">The font size for the header</param>
        /// <exception cref="NotSupportedException">Style-sheets inside the header not supported</exception>
        public void SetHeader<TValue>(in TValue value, TextStyle textStyle, string fontName, FontSize fontSize)
            where TValue : notnull
            => SetHeader(value, textStyle, fontName, FontHelper.GetFontSize(fontSize));

        /// <summary>
        /// Set the given content with the given style and font as header
        /// </summary>
        /// <typeparam name="TValue">The type of the header content</typeparam>
        /// <param name="value">The content for the header</param>
        /// <param name="textStyle">The text style for the header</param>
        /// <param name="fontName">The font (name) for the header</param>
        /// <param name="fontSize">The font size for the header</param>
        /// <exception cref="NotSupportedException">Style-sheets inside the header not supported</exception>
        public void SetHeader<TValue>(in TValue value, TextStyle textStyle, string fontName, double fontSize)
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

            var style = new Style($"MP{MasterStyleCount}", StyleFamily.Header, textStyle, fontName, fontSize);

            StyleHelper.AppendXmlStyleStart(MasterStyle, style);
            StyleHelper.AppendXmlStyle(MasterStyle, style);
            StyleHelper.AppendXmlStyleEnd(MasterStyle);

            // TODO: don't use hard coded name for style informations

            HeaderContent.Append($"<text:p text:style-name=\"MP{MasterStyleCount}\">");
            HeaderContent.Append(value);
            HeaderContent.Append("</text:p>");
        }

        /// <summary>
        /// Set the given content as footer (use stand style and global font)
        /// </summary>
        /// <typeparam name="TValue">The type of the footer content</typeparam>
        /// <param name="value">The content for the footer</param>
        /// <exception cref="NotSupportedException">Style-sheets inside the footer not supported</exception>
        public void SetFooter<TValue>(in TValue value)
            where TValue : notnull
            => SetFooter(value, TextStyle.None, GlobalFontName, GlobalFontSize);

        /// <summary>
        /// Set the given content with the given style as footer (use global font)
        /// </summary>
        /// <typeparam name="TValue">The type of the footer content</typeparam>
        /// <param name="content">The content for the footer</param>
        /// <param name="textStyle">The text style for the footer</param>
        /// <exception cref="NotSupportedException">Style-sheets inside the footer not supported</exception>
        public void SetFooter<TValue>(in TValue content, TextStyle textStyle)
            where TValue : notnull
            => SetFooter(content, textStyle, GlobalFontName, GlobalFontSize);

        /// <summary>
        /// Set the given content with the given style and font as footer
        /// </summary>
        /// <typeparam name="TValue">The type of the footer content</typeparam>
        /// <param name="content">The content for the footer</param>
        /// <param name="textStyle">The text style for the footer</param>
        /// <param name="fontName">The font (name) for the footer</param>
        /// <param name="fontSize">The font size for the footer</param>
        /// <exception cref="NotSupportedException">Style-sheets inside the footer not supported</exception>
        public void SetFooter<TValue>(in TValue content, TextStyle textStyle, string fontName, FontSize fontSize)
            where TValue : notnull
            => SetFooter(content, textStyle, fontName, FontHelper.GetFontSize(fontSize));

        /// <summary>
        /// Set the given content with the given style and font as footer
        /// </summary>
        /// <typeparam name="TValue">The type of the footer content</typeparam>
        /// <param name="content">The content for the footer</param>
        /// <param name="textStyle">The text style for the footer</param>
        /// <param name="fontName">The font (name) for the footer</param>
        /// <param name="fontSize">The font size for the footer</param>
        /// <exception cref="NotSupportedException">Style-sheets inside the footer not supported</exception>
        public void SetFooter<TValue>(in TValue content, TextStyle textStyle, string fontName, double fontSize)
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

            var style = new Style($"MP{MasterStyleCount}", StyleFamily.Footer, textStyle, fontName, fontSize);

            StyleHelper.AppendXmlStyleStart(MasterStyle, style);
            StyleHelper.AppendXmlStyle(MasterStyle, style);
            StyleHelper.AppendXmlStyleEnd(MasterStyle);

            // TODO: don't use hard coded name for style informations

            FooterContent.Append($"<text:p text:style-name=\"MP{MasterStyleCount}\">");
            FooterContent.Append(content);
            FooterContent.Append("</text:p>");
        }
    }
}
