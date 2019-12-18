using NetCoreOdt.Enumerations;
using System.Text;

namespace NetCoreOdt
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // Create and prepare a new ODT document
            using var odtDoucment = new OdtDocument("E:/testTest.odt");

            // Write a single line with a unformatted value to the document
            odtDoucment.Write(long.MinValue);
            odtDoucment.Write(byte.MaxValue);
            odtDoucment.Write(uint.MaxValue);
            odtDoucment.Write(double.NaN);

            // Write a single line with a unformatted text to the document (Note: line breaks "\n" will currently not working)
            odtDoucment.Write("This is a text");

            // Write the content of the given <see cref="StringBuilder"/> into the document (Note: line breaks "\n" will currently not working)
            var stringBuilder = new StringBuilder();
            stringBuilder.Append("This is a text a very very very long text");
            odtDoucment.Write(stringBuilder);

            // Write a single line with a styled value to the document
            odtDoucment.Write(long.MinValue, TextStyle.Bold);
            odtDoucment.Write(byte.MaxValue, TextStyle.Italic);
            odtDoucment.Write(uint.MaxValue, TextStyle.Underline);
            odtDoucment.Write(double.NaN, TextStyle.Bold | TextStyle.Italic | TextStyle.Underline);

            // Write a single line with a unformatted text to the document (Note: line breaks "\n" will currently not working)
            odtDoucment.Write("This is a text", TextStyle.Bold | TextStyle.Underline);

            // Write the content of the given <see cref="StringBuilder"/> into the document (Note: line breaks "\n" will currently not working)
            var stringBuilderFormated = new StringBuilder();
            stringBuilderFormated.Append("This is a text a very very very long text");
            odtDoucment.Write(stringBuilderFormated, TextStyle.Italic | TextStyle.Underline);

            // on Dispose call the ODT document will automatic save and temporary working folder will delete
        }
    }
}
