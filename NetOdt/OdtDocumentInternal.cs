using NetOdt.Constants;
using NetOdt.Enumerations;
using NetOdt.Helper;
using System;
using System.Collections.Generic;
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
        internal XmlDocument ContentFile { get; set; }

        /// <summary>
        /// The raw content before the <see cref="StyleContent"/>
        /// </summary>
        internal StringBuilder BeforeStyleContent { get; }

        /// <summary>
        /// The raw style content
        /// </summary>
        internal StringBuilder StyleContent { get; }

        /// <summary>
        /// The raw content after the <see cref="StyleContent"/> and before the <see cref="TextContent"/>
        /// </summary>
        internal StringBuilder AfterStyleContent { get; }

        /// <summary>
        /// The raw text content
        /// </summary>
        internal StringBuilder TextContent { get; }

        /// <summary>
        /// The raw content after the <see cref="TextContent"/>
        /// </summary>
        internal StringBuilder AfterTextContent { get; }

        /// <summary>
        /// The raw content before the <see cref="ManifestContent"/>
        /// </summary>
        internal StringBuilder BeforeManifestContent { get; }

        /// <summary>
        /// The raw content before the <see cref="MasterStyle"/>
        /// </summary>
        internal StringBuilder BeforeMasterStyleContent { get; }

        /// <summary>
        /// The raw content before the <see cref="HeaderContent"/>
        /// </summary>
        internal StringBuilder BeforeHeaderContent { get; }

        /// <summary>
        /// The raw content after the <see cref="FooterContent"/>
        /// </summary>
        internal StringBuilder AfterFooterContent { get; }

        /// <summary>
        /// The raw manifest content
        /// </summary>
        internal StringBuilder ManifestContent { get; }

        /// <summary>
        /// The uniform resource identifier to the content file (typical inside the <see cref="TempWorkingUri"/>)
        /// </summary>
        internal Uri ContentFileUri { get; }

        /// <summary>
        /// The uniform resource identifier to the manifest file (typical inside the <see cref="TempWorkingUri"/>)
        /// </summary>
        internal Uri ManifestFileUri { get; }

        /// <summary>
        /// The uniform resource identifier to the style file (typical inside the <see cref="TempWorkingUri"/>)
        /// </summary>
        internal Uri StyleFileUri { get; }

        /// <summary>
        /// List that contains all need styles and style names
        /// </summary>
        internal IDictionary<TextStyle, string> NeededStyles { get; }

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

            using(var fileStream = File.OpenRead(StyleFileUri.AbsolutePath))
            {
                using(var textReader = new StreamReader(fileStream))
                {
                    var rawFileContent    = textReader.ReadToEnd();

                    var officeStyleContentSplit = rawFileContent.Split(new string[] { "<office:automatic-styles>" }, StringSplitOptions.None);

                    BeforeMasterStyleContent.Append(officeStyleContentSplit.FirstOrDefault() ?? string.Empty);

                    var automaticStyleContentSplit
                        = officeStyleContentSplit.LastOrDefault().Split(new string[] { "<style:page-layout style:name=\"Mpm1\">" }, StringSplitOptions.None);

                    var styleContentSplit = automaticStyleContentSplit.LastOrDefault().Split(new string[] { "<style:header/><style:footer/>" }, StringSplitOptions.None);

                    BeforeHeaderContent.Append(styleContentSplit.FirstOrDefault() ?? string.Empty);
                    AfterFooterContent.Append(styleContentSplit.LastOrDefault() ?? string.Empty);
                }
            }

        }

        /// <summary>
        /// Write the complete content to the content file (overwrite the existing file)
        /// </summary>
        internal void WriteContent()
        {
            StyleHelper.AddStandardTextStyles(this, NeededStyles);

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

            using(var fileStream = File.Create(StyleFileUri.AbsolutePath))
            {
                using(var textWriter = new StreamWriter(fileStream))
                {
                    textWriter.Write(BeforeMasterStyleContent);
                    textWriter.Write("<office:automatic-styles>");
                    textWriter.Write(MasterStyle);
                    textWriter.Write("<style:page-layout style:name=\"Mpm1\">");
                    textWriter.Write(BeforeHeaderContent);
                    textWriter.Write("<style:header>");
                    textWriter.Write(HeaderContent);
                    textWriter.Write("</style:header>");
                    textWriter.Write("<style:footer>");
                    textWriter.Write(FooterContent);
                    textWriter.Write("</style:footer>");
                    textWriter.Write(AfterFooterContent);
                }
            }
        }

#pragma warning restore IDE0063 // don't use simple using syntax to avoid possible not closed and disposed streams

        /// <summary>
        /// Try to add a new style to the style list and return a style name for the style
        /// </summary>
        /// <param name="style">The style to add to the style list</param>
        internal string TryToAddStyle(TextStyle style)
        {
            string styleName;

            if(NeededStyles.ContainsKey(style))
            {
                NeededStyles.TryGetValue(style, out styleName);
            }
            else
            {
                CheckAcceptStyles(style);
                StyleCount++;
                styleName = $"P{StyleCount}";
                NeededStyles.Add(style, styleName);
            }

            return styleName;
        }

        /// <summary>
        /// Check if the style combination is possible
        /// </summary>
        /// <param name="style">The style to check</param>
        internal void CheckAcceptStyles(TextStyle style)
        {
            switch(style)
            {
                case TextStyle.Left | TextStyle.Center:
                case TextStyle.Left | TextStyle.Right:
                case TextStyle.Left | TextStyle.Justify:
                case TextStyle.Center | TextStyle.Right:
                case TextStyle.Center | TextStyle.Justify:
                case TextStyle.Right | TextStyle.Justify:
                case TextStyle.Left | TextStyle.Center | TextStyle.Right:
                case TextStyle.Left | TextStyle.Center | TextStyle.Justify:
                case TextStyle.Left | TextStyle.Right | TextStyle.Justify:
                case TextStyle.Center | TextStyle.Right | TextStyle.Justify:
                case TextStyle.Left | TextStyle.Center | TextStyle.Right | TextStyle.Justify:
                    throw new ArgumentOutOfRangeException(nameof(style), style, "Only one alignment can be used at same time");

                case TextStyle.Subscript | TextStyle.Superscript:
                    throw new ArgumentOutOfRangeException(nameof(style), style, "Superscript and subscript can't used at same time");

                // TODO: avoid underline combinations (find better way to do this)
            }
        }

        #endregion Internal Methods
    }
}
