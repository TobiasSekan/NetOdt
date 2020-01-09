using System.Text;

namespace NetOdt
{
    /// <summary>
    /// Class to set the global header and footer
    /// </summary>
    public sealed partial class OdtDocument
    {
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
        /// <param name="value">The content  for the header</param>
        public void SetHeader<TValue>(in TValue value)
            where TValue : notnull
        {
            HeaderContent.Append("<text:p>");
            HeaderContent.Append(value);
            HeaderContent.Append("</text:p>");
        }

        /// <summary>
        /// Set the given content as footer
        /// </summary>
        /// <typeparam name="TValue">The type of the footer content</typeparam>
        /// <param name="value">The content for the footer</param>
        public void SetFooter<TValue>(in TValue value)
            where TValue : notnull
        {
            FooterContent.Append("<text:p>");
            FooterContent.Append(value);
            FooterContent.Append("</text:p>");
        }
    }
}
