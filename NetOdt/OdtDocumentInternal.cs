using NetCoreOdt.Helper;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace NetCoreOdt
{
    /// <summary>
    /// Class to create and write ODT documents
    /// </summary>
    public sealed partial class OdtDocument
    {
        #region Internal Properties

        /// <summary>
        /// The XML content of the content file
        ///<para>TODO: #8 - Find a way to use XMLDocument class</para>
        /// </summary>
        internal XmlDocument ContentFile { get; private set; }

        /// <summary>
        /// The raw content before the style content
        /// </summary>
        internal StringBuilder BeforeStyleContent { get; private set; }

        /// <summary>
        /// The raw style content
        /// </summary>
        internal StringBuilder StyleContent { get; private set; }

        /// <summary>
        /// The raw content after the style content and before the text content
        /// </summary>
        internal StringBuilder AfterStyleContent { get; private set; }

        /// <summary>
        /// The raw content after the text content
        /// </summary>
        internal StringBuilder TextContent { get; private set; }

        /// <summary>
        /// The raw text content
        /// </summary>
        internal StringBuilder AfterTextContent { get; private set; }

        /// <summary>
        /// The path to the content file (typical inside the <see cref="TempWorkingPath"/>)
        /// </summary>
        internal string ContentFilePath { get; private set; }

        #endregion Internal Properties

        #region Internal Methods

#pragma warning disable IDE0063 // don't use simple using syntax to avoid possible not closed and disposed streams

        /// <summary>
        /// Read the complete content of the content file
        /// </summary>
        internal void ReadContent()
        {
            // TODO: #8 - Find a way to use XMLDocument class
            ContentFile.Load(ContentFilePath);

            // don't use simple using syntax to avoid possible not closed and disposed streams

            using(var fileStream = File.OpenRead(ContentFilePath))
            {

                using(var textReader = new StreamReader(fileStream))
                {
                    var rawFileContent    = textReader.ReadToEnd();
                    var textContentSplit  = rawFileContent.Split(new string[] { "<text:p text:style-name=\"Standard\"/>" }, StringSplitOptions.None);
                    var styleContentSplit = textContentSplit[0].Split(new string[] { "<office:automatic-styles/>" }, StringSplitOptions.None);

                    BeforeStyleContent.Append(styleContentSplit.ElementAtOrDefault(0) ?? string.Empty);
                    AfterStyleContent.Append(styleContentSplit.ElementAtOrDefault(1) ?? string.Empty);
                    AfterTextContent.Append(textContentSplit.ElementAtOrDefault(1) ?? string.Empty);
                }
            }
        }

        /// <summary>
        /// Write the complete content to the content file (overwrite the existing file)
        /// </summary>
        internal void WriteContent()
        {
            StyleHelper.AddStandardTextStyles(StyleContent);

            // When a document has no tables, we don't need a table style
            if(TableCount > 0)
            {
                StyleHelper.AddTableStyles(StyleContent);
            }

            // don't use simple using syntax to avoid possible not closed and disposed streams

            using(var fileStream = File.Create(ContentFilePath))
            {
                using(var textWriter = new StreamWriter(fileStream))
                {
                    textWriter.Write(BeforeStyleContent);
                    textWriter.Write("<office:automatic-styles>");
                    textWriter.Write(StyleContent);
                    textWriter.Write("</office:automatic-styles>");
                    textWriter.Write(AfterStyleContent);
                    textWriter.Write(TextContent);
                    textWriter.Write(AfterTextContent);
                }
            }
        }

#pragma warning restore IDE0063 // don't use simple using syntax to avoid possible not closed and disposed streams

        #endregion Internal Methods
    }
}
