using NetCoreOdt.Enumerations;
using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml;

namespace NetCoreOdt
{
    /// <summary>
    /// Class to create and write ODT documents
    /// </summary>
    public sealed class OdtDocument : IDisposable
    {
        #region Internal Properties

        /// <summary>
        /// The XML content of the content file
        /// </summary>
        internal XmlDocument ContentFile { get; private set; }

        /// <summary>
        /// The raw content before the text content
        /// </summary>
        internal StringBuilder BeforeTextContent { get; private set; }

        /// <summary>
        /// The raw content after the text content
        /// </summary>
        internal StringBuilder NewTextContent { get; private set; }

        /// <summary>
        /// The raw text content
        /// </summary>
        internal StringBuilder AfteTextContent { get; private set; }

        /// <summary>
        /// The temporary working folder, will delete when <see cref="Dispose"/> is called
        /// </summary>
        internal string TempWorkingPath { get; private set; }

        /// <summary>
        /// The path to the content file (typical inside the <see cref="TempWorkingPath"/>)
        /// </summary>
        internal string ContentFilePath { get; private set; }

        #endregion Internal Properties

        #region Public Constructors

        /// <summary>
        /// Create a new instance of <see cref="OdtDocument"/> and use a automatic generated temporary folder
        /// under the <see cref="Environment.SpecialFolder.LocalApplicationData"/> folder
        /// </summary>
        public OdtDocument()
            : this(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "NetCoreOdt", Guid.NewGuid().ToString()))
        {
        }

        /// <summary>
        /// Create a new instance of <see cref="OdtDocument"/> and use the given temporary folder
        /// </summary>
        /// <param name="tempWorkingPath">The temporary working path for the none zipped document files</param>
        public OdtDocument(string tempWorkingPath)
        {
            TempWorkingPath = tempWorkingPath;
            ContentFilePath = Path.Combine(TempWorkingPath, "content.xml");

            BeforeTextContent = new StringBuilder();
            AfteTextContent   = new StringBuilder();
            NewTextContent    = new StringBuilder();
            ContentFile       = new XmlDocument();

            CreateOdtTemplate();
            ReadContent();
        }

        #endregion Public Constructors

        #region Public Methods - Write

        /// <summary>
        /// Write a single line with a unformatted value to the document
        /// </summary>
        /// <param name="value">The value to write into the document</param>
        public void Write(ValueType value)
        {
            NewTextContent.Append($"<text:p text:style-name=\"Standard\">");
            NewTextContent.Append(value);
            NewTextContent.Append("</text:p>");
        }

        /// <summary>
        /// Write a single line with a unformatted text to the document (Note: line breaks "\n" will currently not working)
        /// </summary>
        /// <param name="text">The text to write into the document</param>
        public void Write(string text)
        {
            NewTextContent.Append($"<text:p text:style-name=\"Standard\">");
            NewTextContent.Append(text);
            NewTextContent.Append("</text:p>");
        }

        /// <summary>
        /// Write the content of the given <see cref="StringBuilder"/> as unformatted text into the document (Note: line breaks "\n" will currently not working)
        /// </summary>
        /// <param name="text">The <see cref="StringBuilder"/> that contains the content for the document</param>
        public void Write(StringBuilder text)
        {
            NewTextContent.Append($"<text:p text:style-name=\"Standard\">");
            NewTextContent.Append(text);
            NewTextContent.Append("</text:p>");
        }

        /// <summary>
        /// Write a single line with a styled value to the document
        /// </summary>
        /// <param name="value">The value to write into the document</param>
        /// <param name="style">The text style of the value</param>
        public void Write(ValueType value, TextStyle style)
        {
            // TODO
            NewTextContent.Append($"<text:p text:style-name=\"Standard\">");
            NewTextContent.Append(value);
            NewTextContent.Append("</text:p>");
        }

        /// <summary>
        /// Write a single line with a styled text to the document (Note: line breaks "\n" will currently not working)
        /// </summary>
        /// <param name="text">The text to write into the document</param>
        /// <param name="style">The text style of the text</param>
        public void Write(string text, TextStyle style)
        {
            // TODO
            NewTextContent.Append($"<text:p text:style-name=\"Standard\">");
            NewTextContent.Append(text);
            NewTextContent.Append("</text:p>");
        }

        /// <summary>
        /// Write the content of the given <see cref="StringBuilder"/> as styled text the document (Note: line breaks "\n" will currently not working)
        /// </summary>
        /// <param name="text">The <see cref="StringBuilder"/> that contains the content for the document</param>
        /// <param name="style">The text style of the content</param>
        public void Write(StringBuilder text, TextStyle style)
        {
            // TODO
            NewTextContent.Append($"<text:p text:style-name=\"Standard\">");
            NewTextContent.Append(text);
            NewTextContent.Append("</text:p>");
        }

        #endregion Public Methods  - Write

        #region Public Methods

        /// <summary>
        /// Save the change content and create the ODT document into the given path , automatic override a existing file
        /// </summary>
        /// <param name="filePath">The save path for the ODT document</param>
        public void Save(string filePath)
            => Save(filePath, overrideExistingFile: true);

        /// <summary>
        /// Save the change content and create the ODT document into the given path 
        /// </summary>
        /// <param name="filePath">The save path for the ODT document</param>
        /// <param name="overrideExistingFile">Indicate that a existing file will be override</param>
        public void Save(string filePath, bool overrideExistingFile)
        {
            WriteContent();

            if(overrideExistingFile && File.Exists(filePath))
            {
                File.Delete(filePath);
            }

            ZipFile.CreateFromDirectory(TempWorkingPath, filePath);
        }

        /// <summary>
        /// Delete the <see cref="TempWorkingPath"/> folder and free all resources
        /// </summary>
        public void Dispose()
        {
            Directory.Delete(TempWorkingPath, true);

            BeforeTextContent.Clear();
            NewTextContent.Clear();
            AfteTextContent.Clear();

            ContentFile     = new XmlDocument();
            TempWorkingPath = string.Empty;
            ContentFilePath = string.Empty;
        }

        #endregion Public Methods

        #region Internal Methods

        /// <summary>
        /// Create a folder with a minimum of files that are need by a ODT file
        /// </summary>
        internal void CreateOdtTemplate()
        {
            var assemblyFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            if(assemblyFolder is null)
            {
                throw new DirectoryNotFoundException($"Assembly directory [{assemblyFolder}] not found");
            }

            var originalFolder = Path.Combine(assemblyFolder, "Original");

            Directory.CreateDirectory(Path.Combine(TempWorkingPath, "Configurations2"));
            Directory.CreateDirectory(Path.Combine(TempWorkingPath, "META-INF"));
            Directory.CreateDirectory(Path.Combine(TempWorkingPath, "Thumbnails"));

            foreach(var file in Directory.GetFiles(originalFolder))
            {
                File.Copy(file, Path.Combine(TempWorkingPath, Path.GetFileName(file)), true);
            }

            // Important: respect the uppercase letters in the folder name
            File.Copy(Path.Combine(originalFolder, "META-INF", "manifest.xml"), Path.Combine(TempWorkingPath, "META-INF", "manifest.xml"), true);

            // Important: respect the uppercase letter in the folder name
            File.Copy(Path.Combine(originalFolder, "Thumbnails", "thumbnail.png"), Path.Combine(TempWorkingPath, "Thumbnails", "thumbnail.png"), true);
        }

        /// <summary>
        /// Read the complete content for the content file
        /// </summary>
        internal void ReadContent()
        {
            ContentFile.Load(ContentFilePath);

            // don't use short using syntax to avoid not closed and disposed stream

            using(var fileStream = File.OpenRead(ContentFilePath))
            {

                using(var textReader = new StreamReader(fileStream))
                {
                    var rawFileContent = textReader.ReadToEnd();
                    var contentSplit   = rawFileContent.Split("<text:p text:style-name=\"Standard\"/>");

                    BeforeTextContent.Append(contentSplit.ElementAtOrDefault(0) ?? string.Empty);
                    AfteTextContent.Append(contentSplit.ElementAtOrDefault(1) ?? string.Empty);
                }
            }

        }

        /// <summary>
        /// Write the complete content to the content file (overwrite the existing file)
        /// </summary>
        internal void WriteContent()
        {
            // don't use short using syntax to avoid not closed and disposed stream

            using(var fileStream = File.Create(ContentFilePath))
            {
                using(var textWriter = new StreamWriter(fileStream))
                {
                    textWriter.Write(BeforeTextContent);
                    textWriter.Write(NewTextContent);
                    textWriter.Write(AfteTextContent);
                }
            }
        }

        #endregion Internal Methods
    }
}
