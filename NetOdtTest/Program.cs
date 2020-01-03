using NetCoreOdt;
using NetCoreOdt.Enumerations;
using NetOdt.Enumerations;
using System;
using System.Data;
using System.Text;

namespace NetOdtTest
{
    public static class Program
    {
        public static void Main(string[] args)
        {
            var uri = new Uri("E:/testTest.odt");

            using var odtDoucment = new OdtDocument("E:/testTest.odt");

            odtDoucment.Append("My Test document", HeaderStyle.Title);

            odtDoucment.AppendTable(GetTable());

            // TODO: need changes inside the "/MATA-INF/mainfest.xml"
            odtDoucment.AppendImage("E:/picture1.jpg", 14.9, 9.8);

            odtDoucment.Append("Unformatted", HeaderStyle.HeadingLevel01);

            odtDoucment.Append(long.MinValue);

            odtDoucment.AppendEmptyLines(countOfEmptyLines: 5);
            odtDoucment.Append(byte.MaxValue);

            odtDoucment.AppendEmptyLines(countOfEmptyLines: 2);
            odtDoucment.Append(uint.MaxValue);

            odtDoucment.AppendEmptyLines(countOfEmptyLines: 1);
            odtDoucment.Append(double.NaN);

            odtDoucment.Append("This is a text");

            var content = new StringBuilder();
            content.Append("This is a text a very very very long text");
            odtDoucment.Append(content);

            odtDoucment.AppendTable(GetTable());

            odtDoucment.Append("Formatted", HeaderStyle.HeadingLevel01);

            // Write a single line with a styled value to the document
            odtDoucment.Append(long.MinValue, TextStyle.Bold);
            odtDoucment.Append(byte.MaxValue, TextStyle.Italic);
            odtDoucment.Append(uint.MaxValue, TextStyle.Underline);
            odtDoucment.Append(double.NaN, TextStyle.Bold | TextStyle.Italic | TextStyle.Underline);

            // Write a single line with a unformatted text to the document (Note: line breaks "\n" will currently not working)
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
