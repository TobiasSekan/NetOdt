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
        internal static void CreateOdtTemplate(string tempWorkingPath)
        {
            var assemblyFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            if(assemblyFolder is null)
            {
                throw new DirectoryNotFoundException($"Assembly directory [{assemblyFolder}] not found");
            }

            var originalFolder = Path.Combine(assemblyFolder, "Original");

            Directory.CreateDirectory(Path.Combine(tempWorkingPath, "Configurations2"));
            Directory.CreateDirectory(Path.Combine(tempWorkingPath, "META-INF"));
            Directory.CreateDirectory(Path.Combine(tempWorkingPath, "Thumbnails"));

            foreach(var file in Directory.GetFiles(originalFolder))
            {
                File.Copy(file, Path.Combine(tempWorkingPath, Path.GetFileName(file)), true);
            }

            // Important: respect the uppercase letters in the folder name
            File.Copy(Path.Combine(originalFolder, "META-INF", "manifest.xml"), Path.Combine(tempWorkingPath, "META-INF", "manifest.xml"), true);

            // Important: respect the uppercase letter in the folder name
            File.Copy(Path.Combine(originalFolder, "Thumbnails", "thumbnail.png"), Path.Combine(tempWorkingPath, "Thumbnails", "thumbnail.png"), true);
        }
    }
}
