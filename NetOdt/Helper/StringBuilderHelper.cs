using System.Text;

namespace NetOdt.Helper
{
    /// <summary>
    /// Helper class to easier work with <see cref="StringBuilder"/> objects
    /// </summary>
    internal static class StringBuilderHelper
    {
        /// <summary>
        /// Test if a given <see cref="StringBuilder"/> content contains line breaks
        /// </summary>
        /// <returns><see langword="true"/> if the content contains line breaks, otherwise <see langword="false"/></returns>
        internal static bool ContainsLineBreaks(in StringBuilder content)
        {
            for(var index = 0; index < content.Length; index++)
            {
                if(content[index] == '\n')
                {
                    return true;
                }
            }

            return false;
        }
    }
}
