using NetOdt.Enumerations;
using NetOdt.Helper;
using NUnit.Framework;
using System.Text;

namespace NetOdtTest.HelperTest
{
    [TestFixture]
    public class StyleHelperTest
    {
        [TestCase("TestName1", StyleFamily.Graphic,     "<style:style style:name=\"TestName1\" style:family=\"graphic\" style:parent-style-name=\"Graphics\">")]
        [TestCase("TestName2", StyleFamily.Paragraph,   "<style:style style:name=\"TestName2\" style:family=\"paragraph\" style:parent-style-name=\"Standard\">")]
        [TestCase("TestName3", StyleFamily.Table,       "<style:style style:name=\"TestName3\" style:family=\"table\">")]
        [TestCase("TestName4", StyleFamily.TableCell,   "<style:style style:name=\"TestName4\" style:family=\"table-cell\">")]
        [TestCase("TestName5", StyleFamily.TableColumn, "<style:style style:name=\"TestName5\" style:family=\"table-column\">")]
        public void AppendXmlStyleStartTest(string styleName, StyleFamily styleFamily, string result)
        {
            var styleContent = new StringBuilder();

            StyleHelper.AppendXmlStyleStart(styleContent, styleName, styleFamily);

            Assert.That(styleContent.ToString(), Is.EqualTo(result));
        }

        [TestCase(TextStyle.Normal,    "")]
        [TestCase(TextStyle.Bold,      "<style:text-properties fo:font-weight=\"bold\" style:font-weight-asian=\"bold\" style:font-weight-complex=\"bold\"/>")]
        [TestCase(TextStyle.Italic,    "<style:text-properties fo:font-style=\"italic\" style:font-style-asian=\"italic\" style:font-style-complex=\"italic\"/>")]
        [TestCase(TextStyle.Underline, "<style:text-properties style:text-underline-style=\"solid\" style:text-underline-width=\"auto\" style:text-underline-color=\"font-color\"/>")]

        [TestCase(TextStyle.Bold | TextStyle.Italic,      "<style:text-properties fo:font-style=\"italic\" fo:font-weight=\"bold\" style:font-style-asian=\"italic\" style:font-weight-asian=\"bold\" style:font-style-complex=\"italic\" style:font-weight-complex=\"bold\"/>")]
        [TestCase(TextStyle.Bold | TextStyle.Underline,   "<style:text-properties style:text-underline-style=\"solid\" style:text-underline-width=\"auto\" style:text-underline-color=\"font-color\" fo:font-weight=\"bold\" style:font-weight-asian=\"bold\" style:font-weight-complex=\"bold\"/>")]
        [TestCase(TextStyle.Italic | TextStyle.Underline, "<style:text-properties fo:font-style=\"italic\" style:text-underline-style=\"solid\" style:text-underline-width=\"auto\" style:text-underline-color=\"font-color\" style:font-style-asian=\"italic\" style:font-style-complex=\"italic\"/>")]

        [TestCase(TextStyle.Bold | TextStyle.Italic | TextStyle.Underline, "<style:text-properties fo:font-style=\"italic\" style:text-underline-style=\"solid\" style:text-underline-width=\"auto\" style:text-underline-color=\"font-color\" fo:font-weight=\"bold\" style:font-style-asian=\"italic\" style:font-weight-asian=\"bold\" style:font-style-complex=\"italic\" style:font-weight-complex=\"bold\"/>")]

        [TestCase(TextStyle.Left    , "<style:paragraph-properties fo:text-align=\"start\" style:justify-single-word=\"false\"/>")]
        [TestCase(TextStyle.Center  , "<style:paragraph-properties fo:text-align=\"center\" style:justify-single-word=\"false\"/>")]
        [TestCase(TextStyle.Right   , "<style:paragraph-properties fo:text-align=\"end\" style:justify-single-word=\"false\"/>")]
        [TestCase(TextStyle.Justify , "<style:paragraph-properties fo:text-align=\"justify\" style:justify-single-word=\"false\"/>")]

        public void AppendXmlStyleTest(TextStyle style, string result)
        {
            var styleContent = new StringBuilder();

            StyleHelper.AppendXmlStyle(styleContent, style);

            Assert.That(styleContent.ToString(), Is.EqualTo(result));
        }

        [Test]
        public void AppendXmlStyleEndTest()
        {
            var styleContent = new StringBuilder();

            StyleHelper.AppendXmlStyleEnd(styleContent);

            Assert.That(styleContent.ToString(), Is.EqualTo("</style:style>"));
        }

    }
}
