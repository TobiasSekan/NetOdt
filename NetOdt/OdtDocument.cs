using NetOdt.Constants;
using NetOdt.Helper;
using System;
using System.Text;
using System.Xml;

namespace NetOdt
{
    /// <summary>
    /// Class to create and write ODT documents
    /// </summary>
    public sealed partial class OdtDocument
    {
        #region Public Properties

        /// <summary>
        /// The uniform resource identifier for the file
        /// </summary>
        public Uri FileUri { get; private set; }

        /// <summary>
        /// The uniform resource identifier for the temporary working folder working folder, the folder will delete when <see cref="Dispose()"/> is called
        /// </summary>
        public Uri TempWorkingUri { get; }

        /// <summary>
        /// The count of the tables
        /// </summary>
        public byte TableCount { get; private set; }

        /// <summary>
        /// The count of picture
        /// </summary>
        public byte PictureCount { get; private set; }

        #endregion Public Properties

        #region Public Constructors

        /// <summary>
        /// Create a new ODT document, save the ODT document as "Unknown.odt" and use a automatic generated temporary folder
        /// under the <see cref="Environment.SpecialFolder.LocalApplicationData"/> folder
        /// </summary>
        public OdtDocument()
            : this(new Uri(FileName.UnkownOdtFile))
        {
        }

        /// <summary>
        /// Create a new ODT document, save the ODT document into the given file path and use a automatic generated temporary folder
        /// under the <see cref="Environment.SpecialFolder.LocalApplicationData"/> folder
        /// </summary>
        /// <param name="filePath">The save path for the ODT document</param>
        public OdtDocument(in string filePath)
            : this(new Uri(filePath))
        {
        }

        /// <summary>
        /// Create a new ODT document, save the ODT document into the given file path and use the given temporary folder
        /// </summary>
        /// <param name="filePath">The save path for the ODT document</param>
        /// <param name="tempWorkingPath">The temporary working path for the none zipped document files</param>
        public OdtDocument(in string filePath, in string tempWorkingPath)
            : this(new Uri(filePath), new Uri(tempWorkingPath))
        {
        }

        /// <summary>
        /// Create a new ODT document, save the ODT document into the given uniform resource identifier and use a automatic generated temporary folder
        /// under the <see cref="Environment.SpecialFolder.LocalApplicationData"/> folder
        /// </summary>
        /// <param name="fileUri">The uniform resource identifier for the ODT document</param>
        public OdtDocument(in Uri fileUri)
            : this(fileUri, UriHelper.Combine(FolderResource.TemporaryRootFolderPath, Guid.NewGuid().ToString()))
        {
        }

        /// <summary>
        /// Create a new ODT document, save the ODT document into the given uniform resource identifier and use the given temporary folder on the uniform resource identifier
        /// </summary>
        /// <param name="fileUri">The uniform resource identifier for the ODT document</param>
        /// <param name="tempWorkingUri">The uniform resource identifier  for the temporary working folder for the none zipped document files</param>
        public OdtDocument(in Uri fileUri, in Uri tempWorkingUri)
        {
            TableCount            = 0;
            PictureCount          = 0;
                                 
            FileUri               = fileUri;
            TempWorkingUri        = tempWorkingUri;
                                 
            ContentFileUri        = UriHelper.Combine(TempWorkingUri, FileName.ContentFile);
            ManifestFileUri       = UriHelper.Combine(TempWorkingUri,
                                                      FolderResource.MainfestFolderName,
                                                      FileName.ManifestFile);

            ContentFile           = new XmlDocument();

            BeforeStyleContent    = new StringBuilder();
            StyleContent          = new StringBuilder();
            AfterStyleContent     = new StringBuilder();
            TextContent           = new StringBuilder();
            AfterTextContent      = new StringBuilder();
            BeforeManifestContent = new StringBuilder();
            ManifestContent       = new StringBuilder();

            OdtDocumentHelper.CreateOdtTemplate(TempWorkingUri);
            ReadContent();
        }

        #endregion Public Constructors
    }
}
