using NetOdt;
using NetOdt.Enumerations;
using NUnit.Framework;
using System;
using System.Data;
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

            odtDocument.SetGlobalFont("Arial", FontSize.Size11);

            odtDocument.SetHeader("Extend documentation", TextStyle.Center | TextStyle.Bold);
            odtDocument.SetFooter("Page 2/5", TextStyle.Right | TextStyle.Italic);

            odtDocument.AppendLine("My Test document", TextStyle.Title);

            odtDocument.AppendTable(GetTable());

            odtDocument.AppendImage("E:/picture1.jpg", width: 10.5, height: 8.0);

            odtDocument.AppendLine("Unformatted", TextStyle.HeadingLevel01);

            odtDocument.AppendEmptyLines(countOfEmptyLines: 5);
            odtDocument.AppendLine(byte.MaxValue, TextStyle.Center | TextStyle.Bold);

            odtDocument.AppendEmptyLines(countOfEmptyLines: 2);
            odtDocument.AppendLine(uint.MaxValue, TextStyle.Right | TextStyle.Italic);

            odtDocument.AppendEmptyLines(countOfEmptyLines: 1);
            odtDocument.AppendLine(double.NaN);

            odtDocument.AppendLine("This is a text");

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

            table.Columns.Add("Dosage", typeof(int));
            table.Columns.Add("Drug", typeof(string));
            table.Columns.Add("Patient", typeof(string));
            table.Columns.Add("Date", typeof(DateTime));

            table.Rows.Add(25, "Indocin\nIndocin", "David", DateTime.Now);
            table.Rows.Add(50, "Enebrel\nEnebrel", "Sam", DateTime.Now);
            table.Rows.Add(10, "Hydralazine\nHydralazine", "Christoff", DateTime.Now);
            table.Rows.Add(21, "Combivent\nCombivent", "Janet", DateTime.Now);
            table.Rows.Add(100, "Dilantin\nDilantin", "Melanie", DateTime.Now);

            return table;
        }
    }
}
