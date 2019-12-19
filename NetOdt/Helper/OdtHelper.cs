using NetOdt.Constants;
using NetOdt.Helper;
using System;
using System.IO;
using System.Reflection;

namespace NetCoreOdt.Helper
{
    /// <summary>
    /// Helper class to easier work with ODT files
    /// </summary>
    internal static class OdtHelper
    {
        /// <summary>
        /// Create a folder with a minimum of files that are need by a ODT document
        /// </summary>
        internal static void CreateOdtTemplate(in Uri tempWorkingUri)
        {
            var assemblyFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            if(assemblyFolder is null)
            {
                throw new DirectoryNotFoundException($"Assembly directory [{assemblyFolder}] not found");
            }

            var originalFolder = Path.Combine(assemblyFolder, "Original");

            DirectoryHelper.CreateDirectory(tempWorkingUri, "Configurations2");
            DirectoryHelper.CreateDirectory(tempWorkingUri, "META-INF");
            DirectoryHelper.CreateDirectory(tempWorkingUri, "Thumbnails");
            DirectoryHelper.CreateDirectory(tempWorkingUri, FolderResource.PictureFolderName);

            foreach(var file in Directory.GetFiles(originalFolder))
            {
                File.Copy(file, Path.Combine(tempWorkingUri.AbsolutePath, Path.GetFileName(file)), true);
            }

            // Important: respect the uppercase letters in the folder name
            File.Copy(Path.Combine(originalFolder, "META-INF", "manifest.xml"), Path.Combine(tempWorkingUri.AbsolutePath, "META-INF", "manifest.xml"), true);

            // Important: respect the uppercase letter in the folder name
            File.Copy(Path.Combine(originalFolder, "Thumbnails", "thumbnail.png"), Path.Combine(tempWorkingUri.AbsolutePath, "Thumbnails", "thumbnail.png"), true);
        }
    }
}
