# Full API description

_More API functions will come in next time, stay tuned_

## Important
**Don't forget to use `using` syntax or call `Dispose()` when finished.**

# Best practice
* Use `AppendEmptyLines(count)` when you need a amount of empty lines
* Use `AppendTable(row, column, text)` when need a pre-filled table
* Use `DataTable` for tables
* Use `StringBuilder` for long text passages
* use `Uri` instead of a `string` for path and file names

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

## Append unformatted values and text into the document

```csharp
// Append a single line with a unformatted value to the document
odtDocument.Append(long.MinValue);
odtDocument.Append(byte.MaxValue);
odtDocument.Append(uint.MaxValue);
odtDocument.Append(double.NaN);

// Append a single line with a unformatted text to the document
odtDocument.Append("This is a text");

// Append the content of the given StringBuilder into the document
var content = new StringBuilder();
content.Append("This is a text a very very very long text");
odtDocument.Append(content);
```

## Append formatted values and text into the document

```csharp
// Append a single line with a styled value to the document
odtDocument.Append(long.MinValue, TextStyle.Bold);
odtDocument.Append(byte.MaxValue, TextStyle.Italic);
odtDocument.Append(uint.MaxValue, TextStyle.Underline);
odtDocument.Append(double.NaN,    TextStyle.Bold | TextStyle.Italic | TextStyle.Underline);

// Append a single line with a unformatted text to the document
odtDocument.Append("This is a text", TextStyle.Bold | TextStyle.Underline);

// Append the content of the given StringBuilder into the document
var content = new StringBuilder();
content.Append("This is a text a very very very long text");
odtDocument.Append(content, TextStyle.Italic | TextStyle.Underline);
```

## Append headers

```csharp
// Append a value with the given header style
odtDocument.Append(long.MinValue, HeaderStyle.HeadingLevel01);
odtDocument.Append(byte.MaxValue, HeaderStyle.HeadingLevel02);
odtDocument.Append(uint.MaxValue, HeaderStyle.HeadingLevel03);
odtDocument.Append(double.NaN,    HeaderStyle.HeadingLevel04);

// Append a value with the given header style
odtDocument.Append("This is a title",  HeaderStyle.Title);

// Append a content with the given header style
var content = new StringBuilder();
content.Append("This is a very very very long footnote");
odtDocument.Append(content, HeaderStyle.Footnote);
```

## Append empty lines

```csharp
/// Append the given count of empty lines
odtDocument.AppendEmptyLines(10);
```

## Append a unformatted table

```csharp
// Append a empty unformatted table with the given row and cell count into the document
odtDocument.AppendTable(3, 3);

// Append a unformatted table with the given row and cell count into the document
// and fill each cell with the given value
odtDocument.AppendTable(3, 3, 0.00);

// Append a unformatted table with the given row and cell count into the document
// and fill each cell with the given text
odtDocument.AppendTable(3, 3, "Fill me");

// Append a unformatted table with the given row and cell count into the document
// and fill each cell with the given content
var content = new StringBuilder();
content.Append("Fill me");
odtDocument.AppendTable(3, 3, content);

// Append a unformatted table and fill it with the given data from the array
var array = new List<List<string>>();
array.Add(new List<string>());
odtDocument.AppendTable(array);

//  Append a unformatted table and fill it with the given data from the DataTable
var table = new DataTable();
odtDocument.AppendTable(dataTable);
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
