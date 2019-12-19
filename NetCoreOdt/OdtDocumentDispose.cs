using System;
using System.IO;
using System.Xml;

namespace NetCoreOdt
{
    /// <summary>
    /// Class to create and write ODT documents
    /// </summary>
    public sealed partial class OdtDocument : IDisposable
    {
        /// <summary>
        /// Save the document (override when existing), delete the <see cref="TempWorkingPath"/> folder and free all resources
        /// </summary>
        public void Dispose()
            => Dispose(overrideExistingFile: true);

        /// <summary>
        /// Save the document , delete the <see cref="TempWorkingPath"/> folder and free all resources
        /// </summary>
        public void Dispose(in bool overrideExistingFile)
        {
            Save(overrideExistingFile);

            Directory.Delete(TempWorkingPath, true);

            BeforeStyleContent.Clear();
            StyleContent.Clear();
            AfterStyleContent.Clear();
            TextContent.Clear();
            AfterTextContent.Clear();

            ContentFile     = new XmlDocument();
            TempWorkingPath = string.Empty;
            ContentFilePath = string.Empty;
        }
    }
}
