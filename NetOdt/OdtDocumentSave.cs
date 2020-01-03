using NetOdt.Helper;
using System;
using System.IO.Compression;

namespace NetCoreOdt
{
    /// <summary>
    /// Class to create and write ODT documents
    /// </summary>
    public sealed partial class OdtDocument : IDisposable
    {
        /// <summary>
        /// Save the change content and create the ODT document into the given path
        /// and automatic override a existing file
        /// </summary>
        /// <param name="filePath">The save path for the ODT document</param>
        public void SaveAs(in string filePath)
            => SaveAs(new Uri(filePath));

        /// <summary>
        /// Save the change content and create the ODT document into the uniform resource identifier
        /// and automatic override a existing file
        /// </summary>
        /// <param name="fileUri">The uniform resource identifier for the ODT document</param>
        public void SaveAs(in Uri fileUri)
            => SaveAs(fileUri, overrideExistingFile: true);

        /// <summary>
        /// Save the change content and create the ODT document into the given path
        /// </summary>
        /// <param name="filePath">The save path for the ODT document</param>
        /// <param name="overrideExistingFile">Indicate that a existing file will be override</param>
        public void SaveAs(in string filePath, in bool overrideExistingFile)
            => SaveAs(new Uri(filePath), overrideExistingFile);

        /// <summary>
        /// Save the change content and create the ODT document into the uniform resource identifier
        /// </summary>
        /// <param name="fileUri">The uniform resource identifier for the ODT document</param>
        /// <param name="overrideExistingFile">Indicate that a existing file will be override</param>
        public void SaveAs(in Uri fileUri, in bool overrideExistingFile)
        {
            FileUri = fileUri;

            Save(overrideExistingFile);
        }

        /// <summary>
        /// Save the change content and create the ODT document and automatic override a existing file
        /// </summary>
        public void Save()
            => Save(overrideExistingFile: true);

        /// <summary>
        /// Save the change content and create the ODT document
        /// </summary>
        /// <param name="overrideExistingFile">Indicate that a existing file will be override</param>
        public void Save(in bool overrideExistingFile)
        {
            WriteContent();

            if(overrideExistingFile && FileHelper.Exists(FileUri))
            {
               FileHelper.Delete(FileUri);
            }

            ZipFile.CreateFromDirectory(TempWorkingUri.AbsolutePath, FileUri.AbsolutePath);
        }
    }
}
