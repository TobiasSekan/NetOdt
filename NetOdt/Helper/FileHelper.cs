using System;
using System.IO;
using System.Runtime.CompilerServices;

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
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static bool Exists(Uri uri)
            => File.Exists(uri.AbsolutePath);

        /// <summary>
        /// Deletes the specified file
        /// </summary>
        /// <param name="uri">The uniform resource identifier of the file to be deleted</param>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static void Delete(Uri uri)
            => File.Delete(uri.AbsolutePath);

        /// <summary>
        /// Deletes the specified file
        /// </summary>
        /// <param name="uriLeft">The left part for the complete path</param>
        /// <param name="pathRight">The right part for the complete path</param>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static void Delete(Uri uriLeft, string pathRight)
            => File.Delete(Path.Combine(uriLeft.AbsolutePath, pathRight));

        /// <summary>
        /// Copies an existing file to a new file
        /// </summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static void Copy(string source, Uri destination)
            => File.Copy(source, destination.AbsolutePath);

        /// <summary>
        /// Copies an existing file to a new file
        /// </summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static void Copy(Uri source, string destination)
            => File.Copy(source.AbsolutePath, destination);

        /// <summary>
        /// Copies an existing file to a new file
        /// </summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static void Copy(Uri source, Uri destination)
            => File.Copy(source.AbsolutePath, destination.AbsolutePath);

        /// <summary>
        /// Return the mine type of the file of the given file path
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        internal static string GetMineType(string path)
        {
            var extension = Path.GetExtension(path).ToLower();

            return extension switch
            {
                ".jpg"  => "image/jpeg",
                ".jpeg" => "image/jpeg",

                _ => throw new ArgumentOutOfRangeException(nameof(path), path, $"mine type of file inside the path not supported")
            };
        }
    }
}
