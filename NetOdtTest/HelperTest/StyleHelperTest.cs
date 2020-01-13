using NetOdt.Class;
using NetOdt.Enumerations;
using NetOdt.Helper;
using NUnit.Framework;
using System.Drawing;
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
            var document = new StringBuilder();
            var style    = new Style(styleName, styleFamily, TextStyle.None, string.Empty, 0.0, Color.Black, Color.Transparent);

            StyleHelper.AppendXmlStyleStart(document, style);

            Assert.That(document.ToString(), Is.EqualTo(result));
        }

        [TestCase(TextStyle.None, "fo:font-weight=\"normal\"")]
        [TestCase(TextStyle.None, "style:font-weight-asian=\"normal\"")]
        [TestCase(TextStyle.None, "style:font-weight-complex=\"normal\"")]
        [TestCase(TextStyle.None, "fo:font-style=\"normal\"")]
        [TestCase(TextStyle.None, "style:font-style-asian=\"normal\"")]
        [TestCase(TextStyle.None, "style:font-style-complex=\"normal\"")]
        [TestCase(TextStyle.None, "style:text-underline-style=\"none\"")]
        [TestCase(TextStyle.None, "style:text-line-through-style=\"none\"")]
        [TestCase(TextStyle.None, "style:text-line-through-type=\"none\"")]

        [TestCase(TextStyle.Bold, "fo:font-weight=\"bold\"")]
        [TestCase(TextStyle.Bold, "style:font-weight-asian=\"bold\"")]
        [TestCase(TextStyle.Bold, "style:font-weight-complex=\"bold\"")]
        [TestCase(TextStyle.Bold, "fo:font-style=\"normal\"")]
        [TestCase(TextStyle.Bold, "style:font-style-asian=\"normal\"")]
        [TestCase(TextStyle.Bold, "style:font-style-complex=\"normal\"")]
        [TestCase(TextStyle.Bold, "style:text-underline-style=\"none\"")]
        [TestCase(TextStyle.Bold, "style:text-line-through-style=\"none\"")]
        [TestCase(TextStyle.Bold, "style:text-line-through-type=\"none\"")]

        [TestCase(TextStyle.Italic, "fo:font-weight=\"normal\"")]
        [TestCase(TextStyle.Italic, "style:font-weight-asian=\"normal\"")]
        [TestCase(TextStyle.Italic, "style:font-weight-complex=\"normal\"")]
        [TestCase(TextStyle.Italic, "fo:font-style=\"italic\"")]
        [TestCase(TextStyle.Italic, "style:font-style-asian=\"italic\"")]
        [TestCase(TextStyle.Italic, "style:font-style-complex=\"italic\"")]
        [TestCase(TextStyle.Italic, "style:text-underline-style=\"none\"")]
        [TestCase(TextStyle.Italic, "style:text-line-through-style=\"none\"")]
        [TestCase(TextStyle.Italic, "style:text-line-through-type=\"none\"")]

        [TestCase(TextStyle.UnderlineSingle, "fo:font-weight=\"normal\"")]
        [TestCase(TextStyle.UnderlineSingle, "style:font-weight-asian=\"normal\"")]
        [TestCase(TextStyle.UnderlineSingle, "style:font-weight-complex=\"normal\"")]
        [TestCase(TextStyle.UnderlineSingle, "fo:font-style=\"normal\"")]
        [TestCase(TextStyle.UnderlineSingle, "style:font-style-asian=\"normal\"")]
        [TestCase(TextStyle.UnderlineSingle, "style:font-style-complex=\"normal\"")]
        [TestCase(TextStyle.UnderlineSingle, "style:text-underline-style=\"solid\"")]
        [TestCase(TextStyle.UnderlineSingle, "style:text-underline-width=\"auto\"")]
        [TestCase(TextStyle.UnderlineSingle, "style:text-underline-color=\"font-color\"")]
        [TestCase(TextStyle.UnderlineSingle, "style:text-line-through-style=\"none\"")]
        [TestCase(TextStyle.UnderlineSingle, "style:text-line-through-type=\"none\"")]

        [TestCase(TextStyle.Stroke, "fo:font-weight=\"normal\"")]
        [TestCase(TextStyle.Stroke, "style:font-weight-asian=\"normal\"")]
        [TestCase(TextStyle.Stroke, "style:font-weight-complex=\"normal\"")]
        [TestCase(TextStyle.Stroke, "fo:font-style=\"normal\"")]
        [TestCase(TextStyle.Stroke, "style:font-style-asian=\"normal\"")]
        [TestCase(TextStyle.Stroke, "style:font-style-complex=\"normal\"")]
        [TestCase(TextStyle.Stroke, "style:text-underline-style=\"none\"")]
        [TestCase(TextStyle.Stroke, "style:text-line-through-style=\"solid\"")]
        [TestCase(TextStyle.Stroke, "style:text-line-through-type=\"single\"")]

        [TestCase(TextStyle.Subscript, "style:text-position=\"sub 58%\"")]

        [TestCase(TextStyle.Superscript, "style:text-position=\"super 58%\"")]

        [TestCase(TextStyle.Bold | TextStyle.Italic, "fo:font-weight=\"bold\"")]
        [TestCase(TextStyle.Bold | TextStyle.Italic, "style:font-weight-asian=\"bold\"")]
        [TestCase(TextStyle.Bold | TextStyle.Italic, "style:font-weight-complex=\"bold\"")]
        [TestCase(TextStyle.Bold | TextStyle.Italic, "fo:font-style=\"italic\"")]
        [TestCase(TextStyle.Bold | TextStyle.Italic, "style:font-style-asian=\"italic\"")]
        [TestCase(TextStyle.Bold | TextStyle.Italic, "style:font-style-complex=\"italic\"")]
        [TestCase(TextStyle.Bold | TextStyle.Italic, "style:text-underline-style=\"none\"")]
        [TestCase(TextStyle.Bold | TextStyle.Italic, "style:text-line-through-style=\"none\"")]
        [TestCase(TextStyle.Bold | TextStyle.Italic, "style:text-line-through-type=\"none\"")]

        [TestCase(TextStyle.Bold | TextStyle.UnderlineSingle, "fo:font-weight=\"bold\"")]
        [TestCase(TextStyle.Bold | TextStyle.UnderlineSingle, "style:font-weight-asian=\"bold\"")]
        [TestCase(TextStyle.Bold | TextStyle.UnderlineSingle, "style:font-weight-complex=\"bold\"")]
        [TestCase(TextStyle.Bold | TextStyle.UnderlineSingle, "fo:font-style=\"normal\"")]
        [TestCase(TextStyle.Bold | TextStyle.UnderlineSingle, "style:font-style-asian=\"normal\"")]
        [TestCase(TextStyle.Bold | TextStyle.UnderlineSingle, "style:font-style-complex=\"normal\"")]
        [TestCase(TextStyle.Bold | TextStyle.UnderlineSingle, "style:text-underline-style=\"solid\"")]
        [TestCase(TextStyle.Bold | TextStyle.UnderlineSingle, "style:text-underline-width=\"auto\"")]
        [TestCase(TextStyle.Bold | TextStyle.UnderlineSingle, "style:text-underline-color=\"font-color\"")]

        [TestCase(TextStyle.Italic | TextStyle.UnderlineSingle, "fo:font-weight=\"normal\"")]
        [TestCase(TextStyle.Italic | TextStyle.UnderlineSingle, "style:font-weight-asian=\"normal\"")]
        [TestCase(TextStyle.Italic | TextStyle.UnderlineSingle, "style:font-weight-complex=\"normal\"")]
        [TestCase(TextStyle.Italic | TextStyle.UnderlineSingle, "fo:font-style=\"italic\"")]
        [TestCase(TextStyle.Italic | TextStyle.UnderlineSingle, "style:font-style-asian=\"italic\"")]
        [TestCase(TextStyle.Italic | TextStyle.UnderlineSingle, "style:font-style-complex=\"italic\"")]
        [TestCase(TextStyle.Italic | TextStyle.UnderlineSingle, "style:text-underline-style=\"solid\"")]
        [TestCase(TextStyle.Italic | TextStyle.UnderlineSingle, "style:text-underline-width=\"auto\"")]
        [TestCase(TextStyle.Italic | TextStyle.UnderlineSingle, "style:text-underline-color=\"font-color\"")]

        [TestCase(TextStyle.Bold | TextStyle.Italic | TextStyle.UnderlineSingle, "fo:font-weight=\"bold\"")]
        [TestCase(TextStyle.Bold | TextStyle.Italic | TextStyle.UnderlineSingle, "style:font-weight-asian=\"bold\"")]
        [TestCase(TextStyle.Bold | TextStyle.Italic | TextStyle.UnderlineSingle, "style:font-weight-complex=\"bold\"")]
        [TestCase(TextStyle.Bold | TextStyle.Italic | TextStyle.UnderlineSingle, "fo:font-style=\"italic\"")]
        [TestCase(TextStyle.Bold | TextStyle.Italic | TextStyle.UnderlineSingle, "style:font-style-asian=\"italic\"")]
        [TestCase(TextStyle.Bold | TextStyle.Italic | TextStyle.UnderlineSingle, "style:font-style-complex=\"italic\"")]
        [TestCase(TextStyle.Bold | TextStyle.Italic | TextStyle.UnderlineSingle, "style:text-underline-style=\"solid\"")]
        [TestCase(TextStyle.Bold | TextStyle.Italic | TextStyle.UnderlineSingle, "style:text-underline-width=\"auto\"")]
        [TestCase(TextStyle.Bold | TextStyle.Italic | TextStyle.UnderlineSingle, "style:text-underline-color=\"font-color\"")]

        // TODO: Update tests
        //
        //[TestCase(TextStyle.Left    , "<style:paragraph-properties fo:text-align=\"start\" style:justify-single-word=\"false\"")]
        //[TestCase(TextStyle.Center  , "<style:paragraph-properties fo:text-align=\"center\" style:justify-single-word=\"false\"")]
        //[TestCase(TextStyle.Right   , "<style:paragraph-properties fo:text-align=\"end\" style:justify-single-word=\"false\"")]
        //[TestCase(TextStyle.Justify , "<style:paragraph-properties fo:text-align=\"justify\" style:justify-single-word=\"false\"")]

        public void AppendXmlStyleTest(TextStyle textStyle, string result)
        {
            var styleContainer = new StringBuilder();
            var style          = new Style(string.Empty, StyleFamily.Paragraph, textStyle, string.Empty, 0.0, Color.Black, Color.Transparent);

            StyleHelper.AppendXmlStyle(styleContainer, style);

            Assert.That(styleContainer.ToString(), Does.StartWith("<style:text-properties"));
            Assert.That(styleContainer.ToString(), Does.EndWith("/>"));

            // Make sure that every property is has a separated with a space
            Assert.That(styleContainer.ToString(), Does.Contain($" {result}"));
        }

        [Test]
        public void AppendXmlStyleEndTest()
        {
            var styleContainer = new StringBuilder();

            StyleHelper.AppendXmlStyleEnd(styleContainer);

            Assert.That(styleContainer.ToString(), Is.EqualTo("</style:style>"));
        }

    }
}
