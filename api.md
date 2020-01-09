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

## Needed imports
```csharp
using NetOdt;               // for all functions
using NetOdt.Enumerations;  // for all styles
```

## Create a document

```csharp
// Create a new ODT document, save the ODT document as "Unknown.odt"
// and use a automatic generated temporary folder
using var odtDocument = new OdtDocument();

// Create a new ODT document, save the ODT document into the given file path
// and use a automatic generated temporary folder
using var odtDocument = new OdtDocument(documentPath: "E:/MyDocument.odt");

// Create a new ODT document, save the ODT document into the given uniform resource identifier
// and use a automatic generated temporary folder
var documentUri  = new Uri("E:/MyDocument.odt")
using var odtDocument = new OdtDocument(documentUri);

// Create a new ODT document, save the ODT document into the given file path
// and use the given temporary folder
using var odtDocument = new OdtDocument(documentPath: "E:/MyDocument.odt", tempPath: "E:/TempFolder");

// Create a new ODT document, save the ODT document into the given uniform resource identifier
// and use the given uniform resource identifier for the temporary folder
var documentUri  = new Uri("E:/MyDocument.odt")
var temporaryUri = new Uri("E:/TempFolder")
using var odtDocument = new OdtDocument(documentUri, temporaryUri);
```

## Set font and font size for the document

```csharp
// Set the font (name) and font size for the document
odtDocument.SetGlobalFont(fontName: "Liberation Serif", fontSize: 12);

// Set the font (name) and font size for the document
odtDocument.SetGlobalFont(fontName: "Liberation Serif", fontSize: FontSize.Size10);
```

The enumeration `NetOdt.Enumerations.FontSize` contains all typical font sizes.

## Header and Footer

```
// Set the given content as header
odtDocument.SetHeader("My header");

// Set the given content with the given style as header
odtDocument.SetHeader("My header", TextStyle.Center | TextStyle.Bold);

// Set the given content as footer
odtDocument.SetFooter("My footer");

// Set the given content with the given style as footer
odtDocument.SetFooter("My footer", TextStyle.Right | TextStyle.Italic);
```

## Append unformatted values and text into the document

```csharp
// Append a single line with a unformatted value to the document
odtDocument.AppendLine(long.MinValue);
odtDocument.AppendLine(byte.MaxValue);
odtDocument.AppendLine(uint.MaxValue);
odtDocument.AppendLine(double.NaN);

// Append a single line with a unformatted text to the document
odtDocument.AppendLine("This is a text");

// Append the content of the given StringBuilder into the document
var content = new StringBuilder();
content.Append("This is a text a very very very long text");
odtDocument.AppendLine(content);
```

## Append formatted values and text into the document

```csharp
// Append a single line with a styled value to the document
odtDocument.AppendLine(long.MinValue, TextStyle.Bold);
odtDocument.AppendLine(byte.MaxValue, TextStyle.Italic);
odtDocument.AppendLine(uint.MaxValue, TextStyle.UnderlineSingle);
odtDocument.AppendLine(double.NaN,    TextStyle.Superscript);

// Append a single line with a unformatted text to the document
odtDocument.AppendLine("This is a stroked text", TextStyle.Stroke);

// Append the content of the given StringBuilder into the document
var content = new StringBuilder();
content.Append("This is a text a very very very long text");
odtDocument.AppendLine(content, TextStyle.Bold | TextStyle.Italic | TextStyle.UnderlineSingle);
```

[All supported styles can find here](./styles.md)

## Append align values and text into the document

```csharp
// Append a single line with a align value to the document
odtDocument.AppendLine(long.MinValue, TextStyle.Left);
odtDocument.AppendLine(uint.MaxValue, TextStyle.Right);

// Append a single line with a unformatted text to the document
odtDocument.AppendLine("This text is centered", TextStyle.Center);

// Append the content of the given StringBuilder into the document
var content = new StringBuilder();
content.Append("This is a text a very very very long text");
odtDocument.AppendLine(content, TextStyle.Justify);
```

[All supported styles can find here](./styles.md)

## Append Tiles, headings, footnotes and other
Note: Don't confuse with **Header** and **Footer**

```csharp
// Append a value with the given style type
odtDocument.AppendLine(long.MinValue, TextStyle.HeadingLevel01);
odtDocument.AppendLine(byte.MaxValue, TextStyle.HeadingLevel02);
odtDocument.AppendLine(uint.MaxValue, TextStyle.HeadingLevel03);
odtDocument.AppendLine(double.NaN,    TextStyle.HeadingLevel04);

// Append a value with the given style type
odtDocument.AppendLine("This is a title",  TextStyle.Title);

// Append a content with the given style type
var content = new StringBuilder();
content.Append("This is a very very very long footnote");
odtDocument.AppendLine(content, TextStyle.Footnote);
```

[All supported styles can find here](./styles.md)

## Append a page break

```charp
/// Append a page break
odtDocument.AppendLines(string.Empty, TextStyle.PageBreak);

/// Append a page break before a bold text passages
odtDocument.AppendLines("This is the last sentence on page one", TextStyle.PageBreak | TextStyle.Bold);
```

## Append empty lines

```csharp
/// Append the given count of empty lines
odtDocument.AppendEmptyLines(10);
```

## Append a unformatted table

```csharp
// Append a empty unformatted table with the given row and cell count into the document
odtDocument.AppendTable(rowCount: 3, columnCount: 3);

// Append a unformatted table with the given row and cell count into the document
// and fill each cell with the given value
odtDocument.AppendTable(rowCount: 3, columnCount: 3, content: 0.00);

// Append a unformatted table with the given row and cell count into the document
// and fill each cell with the given text
odtDocument.AppendTable(rowCount: 3, columnCount: 3, content: "Fill me");

// Append a unformatted table with the given row and cell count into the document
// and fill each cell with the given content
var content = new StringBuilder();
content.Append("Fill me");
odtDocument.AppendTable(rowCount: 3, columnCount: 3, content);

// Append a unformatted table and fill it with the given data from the array
var array = new List<List<string>>();
array.Add(new List<string>());
odtDocument.AppendTable(array);

//  Append a unformatted table and fill it with the given data from the DataTable
var table = new DataTable();
odtDocument.AppendTable(dataTable);
```

Examples for `DataTable`  usage can be found here [C# DataTable Examples](https://www.dotnetperls.com/datatable)

## Append a image

```csharp
// Append a image to the document that is located in the path
odtDocument.AppendImage(path: "E:/picture1.jpg", width: 22.5, height: 14.1);

// Append a image to the document that is located in the uniform resource identifier
var uri = new Uri("E:/picture1.jpg");
odtDocument.AppendImage(uri, width: 22.5, height: 14.1);
```

## Manually save the document

```csharp
// Save the change content and create the ODT document
// and automatic override a existing file
odtDocument.Save();

// Save the change content and create the ODT document
odtDocument.Save(overrideExistingFile: false);

// Save the change content and create the ODT document
// into the given path
odtDocument.SaveAs(path: "E:/MyDocument.odt");

// Save the change content and create the ODT document
// into the given uniform resource identifier
var documentUri = new Uri("E:/MyDocument.odt")
odtDocument.SaveAs(documentUri);

// Save the change content and create the ODT document
// into the given path and automatic override a existing file
odtDocument.SaveAs(path: "E:/MyDocument.odt", overrideExistingFile: false);

// Save the change content and create the ODT document
// into the given uniform resource identifier
// and automatic override a existing file
var documentUri = new Uri("E:/MyDocument.odt")
odtDocument.SaveAs(documentUri, overrideExistingFile: false);
```

## Manually call Dispose method

```csharp
// Save the document (override when existing),
// delete the temporary working folder and free all resources
odtDocument.Dispose();

// Save the document, delete the temporary working folder
// and free all resources
odtDocument.Dispose(overrideExistingFile: false);
```
