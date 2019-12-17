namespace NetCoreOdt
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // Create and prepare a new ODT document
            using var odtDoucment = new OdtDocument();

            // Create and prepare a new ODT document inside the given temporary folder
            // using var odtDoucment = new OdtDocument(tempWorkingPath);

            // Start: testing region

            var contentSplit = odtDoucment.RawFileContent.Split("<text:p text:style-name=\"Standard\"/>");

            odtDoucment.RawFileContent  = contentSplit[0] + "<text:p text:style-name=\"Standard\">blablabla</text:p>" + contentSplit[1];


            // End: testing region


            // Save the ODT document into the given path
            odtDoucment.Save("E:\\testTest.odt");

            // ODT document will dispose here automatically (and remove temporary working folder)
        }
    }
}
