using System;
using System.IO;

namespace NetOdt.Helper
{
    /// <summary>
    /// Helper class to easier work with <see cref="Directory"/>
    /// </summary>
    internal static class DirectoryHelper
    {
        /// <summary>
        /// Creates all directories and subdirectories in the specified path unless they already exist.
        /// </summary>
        /// <param name="uriLeft">The left part of the complete path</param>
        /// <param name="pathRight">The right part of the complete path</param>
        internal static void CreateDirectory(Uri uriLeft, string pathRight)
            => Directory.CreateDirectory(Path.Combine(uriLeft.AbsolutePath, pathRight));

    }
}
