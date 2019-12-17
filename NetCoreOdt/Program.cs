namespace NetCoreOdt
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // Create and prepare a new ODT document
            using var odtDoucment = new OdtDocument();


            // Start: testing region

            var contentSplit = odtDoucment.RawFileContent.Split("<text:p text:style-name=\"Standard\"/>");

            odtDoucment.RawFileContent  = contentSplit[0] + "<text:p text:style-name=\"Standard\">blablabla</text:p>" + contentSplit[1];


            // End: testing region


            // Create and prepare a new ODT document
            odtDoucment.Save("E:\\testTest.odt");

            // ODT document will dispose here
            // -> remove temporary working folder
        }
    }
}
