using NetCoreOdt.Helper;
using System;
using System.IO;
using System.Text;
using System.Xml;

namespace NetCoreOdt
{
    /// <summary>
    /// Class to create and write ODT documents
    /// </summary>
    public sealed partial class OdtDocument
    {
        #region Public Properties

        /// <summary>
        /// The full file path (directory + name for the ODT file)
        /// </summary>
        public string FilePath { get; private set; }

        /// <summary>
        /// The temporary working folder, will delete when <see cref="Dispose"/> is called
        /// </summary>
        public string TempWorkingPath { get; private set; }

        /// <summary>
        /// The count of the tables
        /// </summary>
        public int TableCount { get; private set; }

        #endregion Public Properties

        #region Public Constructors

        /// <summary>
        /// Create a new ODT document, save the ODT document as "Unknown.odt" and use a automatic generated temporary folder
        /// under the <see cref="Environment.SpecialFolder.LocalApplicationData"/> folder
        /// </summary>
        public OdtDocument()
            : this("Unknown.odt")
        {
        }

        /// <summary>
        /// Create a new ODT document, save the ODT document into the given file path and use a automatic generated temporary folder
        /// under the <see cref="Environment.SpecialFolder.LocalApplicationData"/> folder
        /// </summary>
        /// <param name="filePath">The save path for the ODT document</param>
        public OdtDocument(in string filePath)
            : this(filePath,
                   Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "NetCoreOdt", Guid.NewGuid().ToString()))
        {
        }

        /// <summary>
        /// Create a new ODT document, save the ODT document into the given file path and use the given temporary folder
        /// </summary>
        /// <param name="filePath">The save path for the ODT document</param>
        /// <param name="tempWorkingPath">The temporary working path for the none zipped document files</param>
        public OdtDocument(in string filePath, in string tempWorkingPath)
        {
            TableCount         = 0;

            FilePath           = filePath;
            TempWorkingPath    = tempWorkingPath;

            ContentFilePath    = Path.Combine(TempWorkingPath, "content.xml");

            ContentFile        = new XmlDocument();

            BeforeStyleContent = new StringBuilder();
            StyleContent       = new StringBuilder();
            AfterStyleContent  = new StringBuilder();
            TextContent        = new StringBuilder();
            AfterTextContent   = new StringBuilder();

            OdtHelper.CreateOdtTemplate(TempWorkingPath);
            ReadContent();
        }

        #endregion Public Constructors
    }
}
