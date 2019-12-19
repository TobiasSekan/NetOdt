namespace NetOdt.Constants
{
    /// <summary>
    /// Constants for file names
    /// </summary>
    internal static class FileName
    {
        /// <summary>
        /// The file name for an ODT document with no name
        /// </summary>
        internal static string UnkownFile { get; } = "Unknown.odt";

        /// <summary>
        /// The file name for an content file of the ODT document
        /// </summary>
        internal static string ContentFile { get; } = "content.xml";
    }
}
