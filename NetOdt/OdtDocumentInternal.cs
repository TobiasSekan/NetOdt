using NetOdt.Constants;
using NetOdt.Helper;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace NetOdt
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
        /// The raw content before the <see cref="StyleContent"/>
        /// </summary>
        internal StringBuilder BeforeStyleContent { get; private set; }

        /// <summary>
        /// The raw style content
        /// </summary>
        internal StringBuilder StyleContent { get; private set; }

        /// <summary>
        /// The raw content after the <see cref="StyleContent"/> and before the <see cref="TextContent"/>
        /// </summary>
        internal StringBuilder AfterStyleContent { get; private set; }

        /// <summary>
        /// The raw text content
        /// </summary>
        internal StringBuilder TextContent { get; private set; }

        /// <summary>
        /// The raw content after the <see cref="TextContent"/>
        /// </summary>
        internal StringBuilder AfterTextContent { get; private set; }

        /// <summary>
        /// The raw content before the <see cref="ManifestContent"/>
        /// </summary>
        internal StringBuilder BeforeManifestContent { get; private set; }

        /// <summary>
        /// The raw manifest content
        /// </summary>
        internal StringBuilder ManifestContent { get; private set; }

        /// <summary>
        /// The uniform resource identifier to the content file (typical inside the <see cref="TempWorkingUri"/>)
        /// </summary>
        internal Uri ContentFileUri { get; private set; }

        /// <summary>
        /// The uniform resource identifier to the manifest file (typical inside the <see cref="TempWorkingUri"/>)
        /// </summary>
        internal Uri ManifestFileUri { get; private set; }

        #endregion Internal Properties

        #region Internal Methods

#pragma warning disable IDE0063 // don't use simple using syntax to avoid possible not closed and disposed streams

        /// <summary>
        /// Read the complete content of the content file
        /// </summary>
        internal void ReadContent()
        {
            // TODO: #8 - Find a way to use XMLDocument class
            ContentFile.Load(ContentFileUri.AbsolutePath);

            // don't use simple using syntax to avoid possible not closed and disposed streams

            using(var fileStream = File.OpenRead(ContentFileUri.AbsolutePath))
            {
                using(var textReader = new StreamReader(fileStream))
                {
                    var rawFileContent    = textReader.ReadToEnd();
                    var textContentSplit  = rawFileContent.Split(new string[] { "<text:p text:style-name=\"Standard\"/>" }, StringSplitOptions.None);
                    var styleContentSplit = textContentSplit[0].Split(new string[] { "<office:automatic-styles/>" }, StringSplitOptions.None);

                    BeforeStyleContent.Append(styleContentSplit.FirstOrDefault() ?? string.Empty);
                    AfterStyleContent.Append(styleContentSplit.LastOrDefault() ?? string.Empty);
                    AfterTextContent.Append(textContentSplit.LastOrDefault() ?? string.Empty);
                }
            }

            using(var fileStream = File.OpenRead(ManifestFileUri.AbsolutePath))
            {
                using(var textReader = new StreamReader(fileStream))
                {
                    var rawFileContent       = textReader.ReadToEnd();
                    var mainfestContentSplit = rawFileContent.Split(new string[] { "</manifest:manifest>" }, StringSplitOptions.None);

                    BeforeManifestContent.Append(mainfestContentSplit.FirstOrDefault() ?? string.Empty);
                }
            }
        }

        /// <summary>
        /// Write the complete content to the content file (overwrite the existing file)
        /// </summary>
        internal void WriteContent()
        {
            // TODO: Add standard text styles only when needed
            StyleHelper.AddStandardTextStyles(StyleContent);

            // When a document has no tables, we don't need a table style
            if(TableCount > 0)
            {
                StyleHelper.AddTableStyles(StyleContent);
            }

            // When a document has no pictures, we don't need a picture style
            if(PictureCount > 0)
            {
                StyleHelper.AddPictureStyle(StyleContent);
            }
            else
            {
                DirectoryHelper.Delete(TempWorkingUri, FolderResource.PictureFolderName);
            }

            // don't use simple using syntax to avoid possible not closed and disposed streams

            using(var fileStream = File.Create(ContentFileUri.AbsolutePath))
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

            using(var fileStream = File.Create(ManifestFileUri.AbsolutePath))
            {
                using(var textWriter = new StreamWriter(fileStream))
                {
                    textWriter.Write(BeforeManifestContent);
                    textWriter.Write(ManifestContent);
                    textWriter.Write("</manifest:manifest>");
                }
            }
        }

#pragma warning restore IDE0063 // don't use simple using syntax to avoid possible not closed and disposed streams

        #endregion Internal Methods
    }
}
