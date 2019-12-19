using System;
using System.IO;
using System.IO.Compression;

namespace NetCoreOdt
{
    /// <summary>
    /// Class to create and write ODT documents
    /// </summary>
    public sealed partial class OdtDocument : IDisposable
    {
        /// <summary>
        /// Save the change content and create the ODT document into the given path and automatic override a existing file
        /// </summary>
        /// <param name="filePath">The save path for the ODT document</param>
        public void SaveAs(in string filePath)
        {
            FilePath = filePath;

            Save(overrideExistingFile: true);
        }

        /// <summary>
        /// Save the change content and create the ODT document into the given path
        /// </summary>
        /// <param name="filePath">The save path for the ODT document</param>
        /// <param name="overrideExistingFile">Indicate that a existing file will be override</param>
        public void SaveAs(in string filePath, in bool overrideExistingFile)
        {
            FilePath = filePath;

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

            if(overrideExistingFile && File.Exists(FilePath))
            {
                File.Delete(FilePath);
            }

            ZipFile.CreateFromDirectory(TempWorkingPath, FilePath);
        }
    }
}
