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
            var dataTable = GetTable();

            // Create and prepare a new ODT document
            using var odtDoucment = new OdtDocument("E:/testTest.odt");

            odtDoucment.WriteTable(dataTable);

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

            odtDoucment.WriteTable(dataTable);

            // Write a single line with a styled value to the document
            odtDoucment.Write(long.MinValue, TextStyle.Bold);
            odtDoucment.Write(byte.MaxValue, TextStyle.Italic);
            odtDoucment.Write(uint.MaxValue, TextStyle.Underline);
            odtDoucment.Write(double.NaN, TextStyle.Bold | TextStyle.Italic | TextStyle.Underline);

            // Write a single line with a unformatted text to the document (Note: line breaks "\n" will currently not working)
            odtDoucment.Write("This is a text", TextStyle.Bold | TextStyle.Underline);

            // Write the content of the given <see cref="StringBuilder"/> into the document (Note: line breaks "\n" will currently not working)
            var stringBuilderFormated = new StringBuilder();
            stringBuilderFormated.Append("This is a text a very very very long text");
            odtDoucment.Write(stringBuilderFormated, TextStyle.Italic | TextStyle.Underline);

            odtDoucment.WriteTable(3, 3);

            // on Dispose call the ODT document will automatic save and temporary working folder will delete
        }

        static DataTable GetTable()
        {
            var table = new DataTable();
            table.Columns.Add("Dosage", typeof(int));
            table.Columns.Add("Drug", typeof(string));
            table.Columns.Add("Patient", typeof(string));
            table.Columns.Add("Date", typeof(DateTime));

            // Here we add five DataRows.
            table.Rows.Add(25, "Indocin", "David", DateTime.Now);
            table.Rows.Add(50, "Enebrel", "Sam", DateTime.Now);
            table.Rows.Add(10, "Hydralazine", "Christoff", DateTime.Now);
            table.Rows.Add(21, "Combivent", "Janet", DateTime.Now);
            table.Rows.Add(100, "Dilantin", "Melanie", DateTime.Now);
            return table;
        }
    }
}
