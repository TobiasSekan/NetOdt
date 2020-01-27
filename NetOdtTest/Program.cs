using NetOdt;
using NetOdt.Enumerations;
using NUnit.Framework;
using System;
using System.Data;
using System.Drawing;
using System.Text;

namespace NetOdtTest
{
    [TestFixture]
    public class Program
    {
        [Test]
        public void CompleteTest()
        {
            var uri = new Uri("E:/testTest.odt");

            using var odtDocument = new OdtDocument(uri);

            odtDocument.SetGlobalFont("Arial", FontSize.Size12);
            odtDocument.SetGlobalColors(Color.White, Color.Black);

            odtDocument.SetHeader("Extend documentation", TextStyle.Center | TextStyle.Bold, "Liberation Serif", FontSize.Size32, Color.Pink, Color.DarkGray);
            odtDocument.SetFooter("Page 2/5", TextStyle.Right | TextStyle.Italic, "Liberation Sans", FontSize.Size20, Color.Blue, Color.LightYellow);

            odtDocument.AppendLine("My Test document", TextStyle.Title, Color.FromArgb(255, 0, 0), Color.FromArgb(0, 126, 126));

            odtDocument.AppendTable(GetTable());

            odtDocument.AppendImage("E:/picture1.jpg", width: 10.5, height: 8.0);

            odtDocument.AppendLine("Unformatted", TextStyle.HeadingLevel01);

            odtDocument.AppendEmptyLines(countOfEmptyLines: 5);
            odtDocument.AppendLine(byte.MaxValue, TextStyle.Center | TextStyle.Bold);

            odtDocument.AppendEmptyLines(countOfEmptyLines: 2);
            odtDocument.AppendLine(uint.MaxValue, TextStyle.Right | TextStyle.Italic);

            odtDocument.AppendEmptyLines(countOfEmptyLines: 1);
            odtDocument.AppendLine(double.NaN, TextStyle.None, "Liberation Serif", FontSize.Size60);

            odtDocument.AppendLine("This is a text", TextStyle.None, "Liberation Sans", FontSize.Size40);

            var content = new StringBuilder();
            content.Append("This is a text a very very very long text");
            content.Append("This is a text a very very very long text");
            content.Append("This is a text a very very very long text");
            content.Append("This is a text a very very very long text");
            odtDocument.AppendLine(content, TextStyle.Justify);

            odtDocument.AppendTable(GetTable());

            odtDocument.AppendLine("Formatted", TextStyle.HeadingLevel04);

            odtDocument.AppendLine(long.MinValue, TextStyle.Bold | TextStyle.Stroke);
            odtDocument.AppendLine(byte.MaxValue, TextStyle.Italic | TextStyle.Subscript);
            odtDocument.AppendLine(uint.MaxValue, TextStyle.UnderlineSingle | TextStyle.Superscript);
            odtDocument.AppendLine(double.NaN, TextStyle.Bold | TextStyle.Italic | TextStyle.UnderlineSingle | TextStyle.Justify | TextStyle.Stroke);

            odtDocument.AppendLine("This\n\n\nis\n\n\na\n\n\ntext", TextStyle.Bold | TextStyle.UnderlineSingle | TextStyle.Superscript);

            var contentTwo = new StringBuilder();
            content.Append("This is a text a\n very\n\n\nvery very\n\n\nlong text");
            odtDocument.AppendLine(content, TextStyle.PageBreak | TextStyle.Italic | TextStyle.UnderlineSingle | TextStyle.Subscript);

            odtDocument.AppendLine("sub-sub-sub-sub", TextStyle.Subtitle);

            odtDocument.AppendTable(3, 3, "Fill me");
            odtDocument.AppendTable(3, 3, 0.00);

            // on Dispose call the ODT document will automatic save and temporary working folder will delete
        }

        internal static DataTable GetTable()
        {
            var table = new DataTable();

            table.Columns.Add("Float", typeof(int));
            table.Columns.Add("Percentage", typeof(float));
            table.Columns.Add("Currency", typeof(decimal));
            table.Columns.Add("Date", typeof(DateTime));
            table.Columns.Add("Time", typeof(TimeSpan));
            table.Columns.Add("Scientific", typeof(double));
            table.Columns.Add("Fraction", typeof(float));
            table.Columns.Add("Boolean", typeof(bool));
            table.Columns.Add("String", typeof(string));
            table.Columns.Add("StringBuilder", typeof(StringBuilder));

            table.Rows.Add(25, 5.5f, 22.27m, DateTime.MinValue, TimeSpan.MinValue, 25.55, 7.8f, false, "Test", new StringBuilder("Test"));
            table.Rows.Add(-25, -5.5f, -22.27m, DateTime.Now, TimeSpan.Zero, -25.55, -7.8f, true, string.Empty, new StringBuilder());
            table.Rows.Add(1000, 5000f, 22000m, DateTime.MaxValue, TimeSpan.MaxValue, 25.55, 7.8f, false, "TestASSASSAAS", new StringBuilder("TestASDSDDASDS"));

            return table;
        }
    }
}
