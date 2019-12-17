using System.Text;

namespace NetCoreOdt
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // Create and prepare a new ODT document
            using var odtDoucment = new OdtDocument();

            // Create and prepare a new ODT document inside the given temporary folder
            // using var odtDoucment = new OdtDocument("E:/tempFolder");

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

            // Save the ODT document into the given path
            odtDoucment.Save("E:/testTest.odt");

            // ODT document will dispose here automatically (and remove temporary working folder)
        }
    }
}
