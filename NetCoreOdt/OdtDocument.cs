using System;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Xml;

namespace NetCoreOdt
{
    /// <summary>
    /// Class to create and write ODT documents
    /// </summary>
    public sealed class OdtDocument : IDisposable
    {
        #region Public Properties

        /// <summary>
        /// The XML content of the content file
        /// </summary>
        public XmlDocument ContentFile { get; private set; }

        /// <summary>
        /// The temporary working folder, will delete when <see cref="Dispose"/> is called
        /// </summary>
        public string TempWorkingPath { get; private set; }

        #endregion Public Properties

        #region Internal Properties

        /// <summary>
        /// The path to the content file (typical inside the <see cref="TempWorkingPath"/>)
        /// </summary>
        internal string ContentFilePath { get; private set; }

        /// <summary>
        /// The raw content of the content file
        /// </summary>
        internal string RawFileContent { get; set; }

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

            RawFileContent  = string.Empty;
            ContentFile     = new XmlDocument();

            CreateOdtTemplate();
            ReadContent();
        }

        #endregion Public Constructors

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

            ContentFile     = new XmlDocument();
            TempWorkingPath = string.Empty;
            RawFileContent  = string.Empty;
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
                    RawFileContent = textReader.ReadToEnd();
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
                    textWriter.Write(RawFileContent);
                }
            }
        }

        #endregion Internal Methods
    }
}
