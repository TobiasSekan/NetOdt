using System;
using System.IO;

namespace NetOdt.Helper
{
    /// <summary>
    /// Helper class to easier work with <see cref="File"/>
    /// </summary>
    internal static class FileHelper
    {
        /// <summary>
        /// Determines whether the specified file exists
        /// </summary>
        /// <param name="uri">The uniform resource identifier of the file to be deleted</param>
        internal static bool Exists(Uri uri)
            => File.Exists(uri.AbsolutePath);

        /// <summary>
        /// Deletes the specified file
        /// </summary>
        /// <param name="uri">The uniform resource identifier of the file to be deleted</param>
        internal static void Delete(Uri uri)
            => File.Delete(uri.AbsolutePath);
    }
}
