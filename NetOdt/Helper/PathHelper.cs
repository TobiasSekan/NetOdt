using System;
using System.IO;
using System.Runtime.CompilerServices;

namespace NetOdt.Helper
{
    /// <summary>
    /// Helper class to easier work with <see cref="Path"/>
    /// </summary>
    internal static class PathHelper
    {
        /// <summary>
        /// Returns the extension of the specified uniform resource identifier
        /// </summary>
        /// <param name="uri">The uniform resource identifier of the file</param>
        /// <returns></returns>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static string GetExtension(Uri uri)
            => Path.GetExtension(uri.AbsolutePath);
    }
}
