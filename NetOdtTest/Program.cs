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

            using var odtDoucment = new OdtDocument(uri);

            odtDoucment.Append("My Test document", HeaderStyle.Title);

            odtDoucment.AppendTable(GetTable());

            odtDoucment.AppendImage("E:/picture1.jpg", width: 10.5, height: 8.0);

            odtDoucment.Append("Unformatted", HeaderStyle.HeadingLevel01);

            odtDoucment.AppendEmptyLines(countOfEmptyLines: 5);
            odtDoucment.Append(byte.MaxValue, TextAlignment.Center);

            odtDoucment.AppendEmptyLines(countOfEmptyLines: 2);
            odtDoucment.Append(uint.MaxValue, TextAlignment.Right);

            odtDoucment.AppendEmptyLines(countOfEmptyLines: 1);
            odtDoucment.Append(double.NaN);

            odtDoucment.Append("This is a text");

            var content = new StringBuilder();
            content.Append("This is a text a very very very long text");
            content.Append("This is a text a very very very long text");
            content.Append("This is a text a very very very long text");
            content.Append("This is a text a very very very long text");
            odtDoucment.Append(content, TextAlignment.Justify);

            odtDoucment.AppendTable(GetTable());

            odtDoucment.Append("Formatted", HeaderStyle.HeadingLevel04);

            odtDoucment.Append(long.MinValue, TextStyle.Bold);
            odtDoucment.Append(byte.MaxValue, TextStyle.Italic);
            odtDoucment.Append(uint.MaxValue, TextStyle.Underline);
            odtDoucment.Append(double.NaN, TextStyle.Bold | TextStyle.Italic | TextStyle.Underline);

            odtDoucment.Append("This\n\n\nis\n\n\na\n\n\ntext", TextStyle.Bold | TextStyle.Underline);

            var contentTwo = new StringBuilder();
            content.Append("This is a text a\n very\n\n\nvery very\n\n\nlong text");
            odtDoucment.Append(content, TextStyle.Italic | TextStyle.Underline);

            odtDoucment.Append("sub-sub-sub-sub", HeaderStyle.Subtitle);

            odtDoucment.AppendTable(3, 3, "Fill me");
            odtDoucment.AppendTable(3, 3, 0.00);

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
