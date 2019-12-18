using NetCoreOdt.Enumerations;
using System;
using System.Data;
using System.Text;

namespace NetCoreOdt
{
    public class Program
    {
        public static void Main(string[] args)
        {
            using var odtDoucment = new OdtDocument("E:/testTest.odt");

            odtDoucment.WriteTable(GetTable());

            odtDoucment.Write(long.MinValue);
            odtDoucment.Write(byte.MaxValue);
            odtDoucment.Write(uint.MaxValue);
            odtDoucment.Write(double.NaN);

            odtDoucment.Write("This is a text");

            var content = new StringBuilder();
            content.Append("This is a text a very very very long text");
            odtDoucment.Write(content);

            odtDoucment.WriteTable(GetTable());

            // Write a single line with a styled value to the document
            odtDoucment.Write(long.MinValue, TextStyle.Bold);
            odtDoucment.Write(byte.MaxValue, TextStyle.Italic);
            odtDoucment.Write(uint.MaxValue, TextStyle.Underline);
            odtDoucment.Write(double.NaN, TextStyle.Bold | TextStyle.Italic | TextStyle.Underline);

            // Write a single line with a unformatted text to the document (Note: line breaks "\n" will currently not working)
            odtDoucment.Write("This\n\n\nis\n\n\na\n\n\ntext", TextStyle.Bold | TextStyle.Underline);

            var contentTwo = new StringBuilder();
            content.Append("This is a text a\n very\n\n\nvery very\n\n\nlong text");
            odtDoucment.Write(content, TextStyle.Italic | TextStyle.Underline);

            odtDoucment.WriteTable(3, 3, "Fill me");
            odtDoucment.WriteTable(3, 3, 0.00);

            // on Dispose call the ODT document will automatic save and temporary working folder will delete
        }

        static DataTable GetTable()
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
