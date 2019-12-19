# Full API description

_More API functions will come in next time, stay tuned_

## Important
**Don't forget to use `using` syntax or call `Dispose()` when finished.**

# Best practice
* Use `WriteEmptyLines(count)` when you need a amount of empty lines
* Use `WriteTable(row, column, text)` when need a pre-filled table
* Use `DataTable` for tables
* Use `StringBuilder` for long text passages
* use `Uri` instead of a `string`for path and file names

## Create a document

```csharp
// Create a new ODT document, save the ODT document as "Unknown.odt"
// and use a automatic generated temporary folder
using var odtDocument = new OdtDocument();

// Create a new ODT document, save the ODT document into the given file path
// and use a automatic generated temporary folder
using var odtDocument = new OdtDocument("E:/MyDocument.odt");

// Create a new ODT document, save the ODT document into the given uniform resource identifier
// and use a automatic generated temporary folder
var documentUri  = new Uri("E:/MyDocument.odt")
using var odtDocument = new OdtDocument(documentUri);

// Create a new ODT document, save the ODT document into the given file path
// and use the given temporary folder
using var odtDocument = new OdtDocument("E:/MyDocument.odt", "E:/TempFolder");

// Create a new ODT document, save the ODT document into the given uniform resource identifier
// and use the given uniform resource identifier for the temporary folder
var documentUri  = new Uri("E:/MyDocument.odt")
var temporaryUri = new Uri("E:/TempFolder")
using var odtDocument = new OdtDocument(documentUri, temporaryUri);
```

## Write unformatted values and text into the document

```csharp
// Write a single line with a unformatted value to the document
odtDocument.Write(long.MinValue);
odtDocument.Write(byte.MaxValue);
odtDocument.Write(uint.MaxValue);
odtDocument.Write(double.NaN);

// Write a single line with a unformatted text to the document
odtDocument.Write("This is a text");

// Write the content of the given StringBuilder into the document
var content = new StringBuilder();
content.Append("This is a text a very very very long text");
odtDocument.Write(content);
```

## Write formatted values and text into the document

```csharp
// Write a single line with a styled value to the document
odtDocument.Write(long.MinValue, TextStyle.Bold);
odtDocument.Write(byte.MaxValue, TextStyle.Italic);
odtDocument.Write(uint.MaxValue, TextStyle.Underline);
odtDocument.Write(double.NaN,    TextStyle.Bold | TextStyle.Italic | TextStyle.Underline);

// Write a single line with a unformatted text to the document
odtDocument.Write("This is a text", TextStyle.Bold | TextStyle.Underline);

// Write the content of the given StringBuilder into the document
var content = new StringBuilder();
content.Append("This is a text a very very very long text");
odtDocument.Write(content, TextStyle.Italic | TextStyle.Underline);
```

## Write headers

```csharp
// Write a value with the given header style
odtDocument.Write(long.MinValue, HeaderStyle.HeadingLevel01);
odtDocument.Write(byte.MaxValue, HeaderStyle.HeadingLevel02);
odtDocument.Write(uint.MaxValue, HeaderStyle.HeadingLevel03);
odtDocument.Write(double.NaN,    HeaderStyle.HeadingLevel04);

// Write a value with the given header style
odtDocument.Write("This is a title",  HeaderStyle.Title);

// Write a content with the given header style
var content = new StringBuilder();
content.Append("This is a very very very long footnote");
odtDocument.Write(content, HeaderStyle.Footnote);
```

## Write empty lines

```csharp
/// Write the given count of empty lines
odtDocument.WriteEmptyLines(10);
```

## Write a unformatted table

```csharp
// Write an empty unformatted table with the given row and cell count into the document
odtDocument.WriteTable(3, 3);

// Write an unformatted table with the given row and cell count into the document
// and fill each cell with the given value
odtDocument.WriteTable(3, 3, 0.00);

// Write an unformatted table with the given row and cell count into the document
// and fill each cell with the given text
odtDocument.WriteTable(3, 3, "Fill me");

// Write an unformatted table with the given row and cell count into the document
// and fill each cell with the given content
var content = new StringBuilder();
content.Append("Fill me");
odtDocument.WriteTable(3, 3, content);

// Write a unformatted table and fill it with the given data from the array
var array = new List<List<string>>();
array.Add(new List<string>());
odtDocument.WriteTable(array);

//  Write a unformatted table and fill it with the given data from the DataTable
var table = new DataTable();
odtDocument.WriteTable(dataTable);
```

Examples for `DataTable`  usage can be found here [C# DataTable Examples](https://www.dotnetperls.com/datatable)

## Manually save the document

```csharp
// Save the change content and create the ODT document and automatic override a existing file
odtDocument.Save();

// Save the change content and create the ODT document
odtDocument.Save(overrideExistingFile: false);

// Save the change content and create the ODT document into the given path
odtDocument.SaveAs("E:/MyDocument.odt");

// Save the change content and create the ODT document into the given uniform resource identifier
var documentUri = new Uri("E:/MyDocument.odt")
odtDocument.SaveAs(documentUri);

// Save the change content and create the ODT document into the given path
// and automatic override a existing file
odtDocument.SaveAs("E:/MyDocument.odt", overrideExistingFile: false);

// Save the change content and create the ODT document into the given uniform resource identifier
// and automatic override a existing file
var documentUri = new Uri("E:/MyDocument.odt")
odtDocument.SaveAs(documentUri, overrideExistingFile: false);
```

## Manually call Dispose method

```csharp
// Save the document (override when existing), delete the temporary working folder
// and free all resources
odtDocument.Dispose();

// Save the document, delete the temporary working folder
// and free all resources
odtDocument.Dispose(overrideExistingFile: false);
```
