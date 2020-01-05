using NetOdt.Enumerations;
using NetOdt.Helper;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace NetCoreOdt
{
    /// <summary>
    /// Class to create and write ODT documents
    /// </summary>
    public sealed partial class OdtDocument : IDisposable
    {
        /// <summary>
        /// Append an empty unformatted table with the given row and cell count into the document
        /// </summary>
        /// <param name="rowCount">The count of the rows</param>
        /// <param name="columnCount">The count of the columns</param>
        public void AppendTable(in int rowCount, in int columnCount)
        {
            TableCount++;

            TextContent.Append($"<table:table table:name=\"Tabelle{TableCount}\" table:style-name=\"Tabelle1\">");
            TextContent.Append($"<table:table-column table:style-name=\"Tabelle1.A\" table:number-columns-repeated=\"{columnCount}\"/>");

            for(var rowNumber = 1; rowNumber <= rowCount; rowNumber++)
            {
                TextContent.Append("<table:table-row>");

                for(var columnNumber = 1; columnNumber <= columnCount; columnNumber++)
                {
                    TextContent.Append($"<table:table-cell table:style-name=\"Tabelle1.{StyleHelper.GetTableCellStyleName(rowNumber, columnNumber, columnCount)}\" office:value-type=\"string\">");
                    TextContent.Append($"</table:table-cell>");
                }

                TextContent.Append("</table:table-row>");
            }

            TextContent.Append("</table:table>");
        }

        /// <summary>
        /// Append a unformatted table with the given row and cell count into the document and fill each cell with the given value
        /// </summary>
        /// <param name="rowCount">The count of the rows</param>
        /// <param name="columnCount">The count of the columns</param>
        /// <param name="value">The value for each cell</param>
        public void AppendTable(in int rowCount, in int columnCount, in ValueType value)
        {
            TableCount++;

            TextContent.Append($"<table:table table:name=\"Tabelle{TableCount}\" table:style-name=\"Tabelle1\">");
            TextContent.Append($"<table:table-column table:style-name=\"Tabelle1.A\" table:number-columns-repeated=\"{columnCount}\"/>");

            for(var rowNumber = 1; rowNumber <= rowCount; rowNumber++)
            {
                TextContent.Append("<table:table-row>");

                for(var columnNumber = 1; columnNumber <= columnCount; columnNumber++)
                {
                    TextContent.Append($"<table:table-cell table:style-name=\"Tabelle1.{StyleHelper.GetTableCellStyleName(rowNumber, columnNumber, columnCount)}\" office:value-type=\"string\">");
                    Append(value);
                    TextContent.Append($"</table:table-cell>");
                }

                TextContent.Append("</table:table-row>");
            }

            TextContent.Append("</table:table>");
        }

        /// <summary>
        /// Append a unformatted table with the given row and cell count into the document and fill each cell with the given text
        /// </summary>
        /// <param name="rowCount">The count of the rows</param>
        /// <param name="columnCount">The count of the columns</param>
        /// <param name="text">The text for each cell</param>
        public void AppendTable(in int rowCount, in int columnCount, in string text)
        {
            TableCount++;

            TextContent.Append($"<table:table table:name=\"Tabelle{TableCount}\" table:style-name=\"Tabelle1\">");
            TextContent.Append($"<table:table-column table:style-name=\"Tabelle1.A\" table:number-columns-repeated=\"{columnCount}\"/>");

            for(var rowNumber = 1; rowNumber <= rowCount; rowNumber++)
            {
                TextContent.Append("<table:table-row>");

                for(var columnNumber = 1; columnNumber <= columnCount; columnNumber++)
                {
                    TextContent.Append($"<table:table-cell table:style-name=\"Tabelle1.{StyleHelper.GetTableCellStyleName(rowNumber, columnNumber, columnCount)}\" office:value-type=\"string\">");
                    Append(text);
                    TextContent.Append($"</table:table-cell>");
                }

                TextContent.Append("</table:table-row>");
            }

            TextContent.Append("</table:table>");
        }

        /// <summary>
        /// Append a unformatted table with the given row and cell count into the document and fill each cell with the given content
        /// </summary>
        /// <param name="rowCount">The count of the rows</param>
        /// <param name="columnCount">The count of the columns</param>
        /// <param name="content">The content for each cell</param>
        public void AppendTable(in int rowCount, in int columnCount, in StringBuilder content)
        {
            TableCount++;

            TextContent.Append($"<table:table table:name=\"Tabelle{TableCount}\" table:style-name=\"Tabelle1\">");
            TextContent.Append($"<table:table-column table:style-name=\"Tabelle1.A\" table:number-columns-repeated=\"{columnCount}\"/>");

            for(var rowNumber = 1; rowNumber <= rowCount; rowNumber++)
            {
                TextContent.Append("<table:table-row>");

                for(var columnNumber = 1; columnNumber <= columnCount; columnNumber++)
                {
                    TextContent.Append($"<table:table-cell table:style-name=\"Tabelle1.{StyleHelper.GetTableCellStyleName(rowNumber, columnNumber, columnCount)}\" office:value-type=\"string\">");
                    Append(content);
                    TextContent.Append($"</table:table-cell>");
                }

                TextContent.Append("</table:table-row>");
            }

            TextContent.Append("</table:table>");
        }

        /// <summary>
        /// Append a unformatted table and fill it with the given data from the value array
        /// </summary>
        /// <param name="valueArray">The array that contains the values for the table</param>
        public void AppendTable(in IEnumerable<IEnumerable<ValueType>> valueArray)
        {
            TableCount++;

            TextContent.Append($"<table:table table:name=\"Tabelle{TableCount}\" table:style-name=\"Tabelle1\">");
            TextContent.Append($"<table:table-column table:style-name=\"Tabelle1.A\" table:number-columns-repeated=\"{valueArray.Count()}\"/>");

            var rowNumber = 0;

            foreach(var dataRow in valueArray)
            {
                rowNumber++;

                if(dataRow is null)
                {
                    continue;
                }

                TextContent.Append("<table:table-row>");

                var columnCount  = dataRow.Count();
                var columnNumber = 0;

                foreach(var column in dataRow)
                {
                    columnNumber++;

                    if(column is null)
                    {
                        continue;
                    }

                    TextContent.Append($"<table:table-cell table:style-name=\"Tabelle1.{StyleHelper.GetTableCellStyleName(rowNumber, columnNumber, columnCount)}\" office:value-type=\"string\">");
                    Append(column, TextStyle.Normal);
                    TextContent.Append($"</table:table-cell>");
                }

                TextContent.Append("</table:table-row>");
            }

            TextContent.Append("</table:table>");
        }

        /// <summary>
        /// Append a unformatted table and fill it with the given data from the string array
        /// </summary>
        /// <param name="stringArray">The array that contains the strings for the table</param>
        public void AppendTable(in IEnumerable<IEnumerable<string>> stringArray)
        {
            TableCount++;

            TextContent.Append($"<table:table table:name=\"Tabelle{TableCount}\" table:style-name=\"Tabelle1\">");
            TextContent.Append($"<table:table-column table:style-name=\"Tabelle1.A\" table:number-columns-repeated=\"{stringArray.Count()}\"/>");

            var rowNumber = 0;

            foreach(var dataRow in stringArray)
            {
                rowNumber++;

                if(dataRow is null)
                {
                    continue;
                }

                TextContent.Append("<table:table-row>");

                var columnCount  = dataRow.Count();
                var columnNumber = 0;

                foreach(var column in dataRow)
                {
                    columnNumber++;

                    if(column is null)
                    {
                        continue;
                    }

                    TextContent.Append($"<table:table-cell table:style-name=\"Tabelle1.{StyleHelper.GetTableCellStyleName(rowNumber, columnNumber, columnCount)}\" office:value-type=\"string\">");
                    Append(column, TextStyle.Normal);
                    TextContent.Append($"</table:table-cell>");
                }

                TextContent.Append("</table:table-row>");
            }

            TextContent.Append("</table:table>");
        }

        /// <summary>
        /// Append a unformatted table and fill it with the given content from the content array
        /// </summary>
        /// <param name="contentArray">The array that contains the contents for the table</param>
        public void AppendTable(in IEnumerable<IEnumerable<StringBuilder>> contentArray)
        {
            TableCount++;

            TextContent.Append($"<table:table table:name=\"Tabelle{TableCount}\" table:style-name=\"Tabelle1\">");
            TextContent.Append($"<table:table-column table:style-name=\"Tabelle1.A\" table:number-columns-repeated=\"{contentArray.Count()}\"/>");

            var rowNumber = 0;

            foreach(var dataRow in contentArray)
            {
                rowNumber++;

                if(dataRow is null)
                {
                    continue;
                }

                TextContent.Append("<table:table-row>");

                var columnCount  = dataRow.Count();
                var columnNumber = 0;

                foreach(var column in dataRow)
                {
                    columnNumber++;

                    if(column is null)
                    {
                        continue;
                    }

                    TextContent.Append($"<table:table-cell table:style-name=\"Tabelle1.{StyleHelper.GetTableCellStyleName(rowNumber, columnNumber, columnCount)}\" office:value-type=\"string\">");
                    Append(column, TextStyle.Normal);
                    TextContent.Append($"</table:table-cell>");
                }

                TextContent.Append("</table:table-row>");
            }

            TextContent.Append("</table:table>");
        }

        /// <summary>
        /// Append a unformatted table and fill it with the given data from the <see cref="DataTable"/>
        /// </summary>
        /// <param name="dataTable">The <see cref="DataTable"/> that contains the data for the table</param>
        public void AppendTable(in DataTable dataTable)
        {
            TableCount++;

            TextContent.Append($"<table:table table:name=\"Tabelle{TableCount}\" table:style-name=\"Tabelle1\">");
            TextContent.Append($"<table:table-column table:style-name=\"Tabelle1.A\" table:number-columns-repeated=\"{dataTable.Columns.Count}\"/>");

            var rowNumber = 0;

            foreach(DataRow? dataRow in dataTable.Rows)
            {
                rowNumber++;

                if(dataRow is null)
                {
                    continue;
                }

                TextContent.Append("<table:table-row>");

                var columnCount  = dataRow.ItemArray.Length;
                var columnNumber = 0;

                foreach(var column in dataRow.ItemArray)
                {
                    columnNumber++;

                    if(column is null)
                    {
                        continue;
                    }

                    TextContent.Append($"<table:table-cell table:style-name=\"Tabelle1.{StyleHelper.GetTableCellStyleName(rowNumber, columnNumber, columnCount)}\" office:value-type=\"string\">");
                    Append(column.ToString() ?? string.Empty, TextStyle.Normal);
                    TextContent.Append($"</table:table-cell>");
                }

                TextContent.Append("</table:table-row>");
            }

            TextContent.Append("</table:table>");
        }
    }
}
