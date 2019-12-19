using System;
using System.IO;

namespace NetOdt.Constants
{
    /// <summary>
    /// Constants for folder resources
    /// </summary>
    internal static class FolderResource
    {
        /// <summary>
        /// The name of the temporary root folder name
        /// </summary>
        internal static string TemporaryRootFolderName { get; } = "NetOdt";

        /// <summary>
        /// The path to the local application data
        /// </summary>
        internal static string LocalApplicationDataPath { get; } = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

        /// <summary>
        /// The name of the temporary root folder path
        /// </summary>
        internal static string TemporaryRootFolderPath { get; } = Path.Combine(LocalApplicationDataPath, TemporaryRootFolderName);

    }
}
