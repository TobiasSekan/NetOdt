using System;
using System.IO;

namespace NetOdt.Helper
{
    /// <summary>
    /// Helper class to easier work with <see cref="Uri"/>
    /// </summary>
    internal static class UriHelper
    {
        /// <summary>
        /// Combine to path and return the resulting <see cref="Uri"/>
        /// </summary>
        /// <param name="pathLeft">The left part for the complete path</param>
        /// <param name="pathRight">The right part for the complete path</param>
        /// <returns>A <see cref="Uri"/> with the complete path</returns>
        internal static Uri Combine(string pathLeft, string pathRight)
            => new Uri(Path.Combine(pathLeft, pathRight));

        /// <summary>
        /// Combine a path and a <see cref="Uri"/> and return the resulting <see cref="Uri"/>
        /// </summary>
        /// <param name="uriLeft">The left part for the complete path</param>
        /// <param name="pathRight">The right part for the complete path</param>
        /// <returns>A <see cref="Uri"/> with the complete path</returns>
        internal static Uri Combine(Uri uriLeft, string pathRight)
            => new Uri(Path.Combine(uriLeft.AbsolutePath, pathRight));

        /// <summary>
        /// Combine to <see cref="Uri"/> and return the resulting <see cref="Uri"/>
        /// </summary>
        /// <param name="uriLeft">The left part for the complete path</param>
        /// <param name="uriRight">The right part for the complete path</param>
        /// <returns>A <see cref="Uri"/> with the complete path</returns>
        internal static Uri Combine(Uri uriLeft, Uri uriRight)
            => new Uri(Path.Combine(uriLeft.AbsolutePath, uriRight.AbsolutePath));
    }
}
