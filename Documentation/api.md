# Full API description

_More API functions will come in next time, stay tuned_

Contents:
* [Important](#important)
* [Best practice](#best-practice)
* [Needed imports](#needed-imports)
* [Create a document](#create-a-document)
* [Set global font and font size for the document](#set-global-font-and-font-size-for-the-document)
* [Set global foreground and background color for the document](#set-global-foreground-and-background-color-for-the-document)
* [Set header and Footer](#set-header-and-footer)
* [Append unformatted values and text into the document](#append-unformatted-values-and-text-into-the-document)
* [Append formatted values and text into the document](#append-formatted-values-and-text-into-the-document)
* [Append align values and text into the document](#append-align-values-and-text-into-the-document)
* [Append style-sheets (tiles, headings, footnotes and other)](#append-style-sheets-tiles-headings-footnotes-and-other)
* [Append a page break](#append-a-page-break)
* [Append empty lines](#append-empty-lines)
* [Append a unformatted table](#append-a-unformatted-table)
* [Append a image](#append-a-image)
* [Manually save the document](#manually-save-the-document)
* [Manually call Dispose method](#manually-call-dispose-method)

## Important
**Don't forget to use `using` syntax or call `Dispose()` when finished.**

## Best practice
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

## Set global font and font size for the document

```csharp
// Set the font (name) and font size for the document
odtDocument.SetGlobalFont(fontName: "Liberation Serif", fontSize: 12);

// Set the font (name) and font size for the document
odtDocument.SetGlobalFont(fontName: "Liberation Serif", fontSize: FontSize.Size10);
```

The enumeration `NetOdt.Enumerations.FontSize` contains all typical font sizes.

## Set global foreground and background color for the document

```csharp
// Set the global foreground and background color for the text passages of the document
odtDocument.SetGlobalColors(foregroundColor: Color.Red, backgroundColor: Color.LightGrey);

// Set the global foreground and background color for the text passages of the document
odtDocument.SetGlobalColors(foregroundColor: Color.FromArgb(0, 0, 0), backgroundColor: Color.Transparent);
```

Note: Alpha-channel is not support and will ignored.

## Set header and Footer

Note: [Don't confuse with heading and footnote](#append-style-sheets-tiles-headings-footnotes-and-other)

```csharp
// Set the given content as header (use standard style and global font)
odtDocument.SetHeader("My header");

// Set the given content with the given style as header (use global font)
odtDocument.SetHeader("My header", TextStyle.Center | TextStyle.Bold);

// Set the given content with the given style and font as header
odtDocument.SetHeader("My header", TextStyle.Center | TextStyle.Bold, "Liberation Sans", FontSize.Size22);

// Set the given content with the given style and color as header
odtDocument.SetHeader("My header", TextStyle.Center | TextStyle.Bold, Color.Gray, Color.Transparent);

// Set the given content with the given style, font and color as header
odtDocument.SetHeader("My header", TextStyle.Center | TextStyle.Bold, "Liberation Sans", FontSize.Size22, Color.Gray, Color.Transparent);

// Set the given content as footer (use standard style and global font)
odtDocument.SetFooter("My footer");

// Set the given content with the given style as footer (use global font)
odtDocument.SetFooter("My footer", TextStyle.Right | TextStyle.Italic);

// Set the given content with the given style and font as footer
odtDocument.SetFooter("My footer", TextStyle.Right | TextStyle.Italic, "Arial", 10.3);

// Set the given content with the given style and color as footer
odtDocument.SetFooter("My footer", TextStyle.Right | TextStyle.Italic, Color.Gray, Color.Transparent);

// Set the given content with the given style, font and color as footer
odtDocument.SetFooter("My footer", TextStyle.Right | TextStyle.Italic, "Arial", 10.3, Color.Gray, Color.Transparent);
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
odtDocument.AppendLine(byte.MaxValue, TextStyle.Italic, "Liberation Serif", 10.9);
odtDocument.AppendLine(uint.MaxValue, TextStyle.UnderlineSingle, "Arial", FontSize.Size10);
odtDocument.AppendLine(double.NaN,    TextStyle.Superscript, Color.Blue, Color.Black);

// Append a single line with a formatted text to the document
odtDocument.AppendLine("This is a stroked text", TextStyle.Stroke, "Arial", FontSize.Size22);
odtDocument.AppendLine("This is a stroked text", TextStyle.Stroke, Color.Blue, Color.Black);
odtDocument.AppendLine("This is a stroked text", TextStyle.Stroke, "Arial", FontSize.Size22, Color.Blue, Color.Black);

// Append the content of the given StringBuilder into the document
var content = new StringBuilder();
content.Append("This is a text a very very very long text");
odtDocument.AppendLine(content, TextStyle.Bold | TextStyle.Italic | TextStyle.UnderlineSingle, "Liberation Sans", 22.7);
odtDocument.AppendLine(content, TextStyle.Bold | TextStyle.Italic | TextStyle.UnderlineSingle, Color.Red, Color.LightGray);
odtDocument.AppendLine(content, TextStyle.Bold | TextStyle.Italic | TextStyle.UnderlineSingle, "Liberation Sans", 22.7, Color.Red, Color.LightGray);
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

## Append style-sheets (tiles, headings, footnotes and other)

Note: [Don't confuse with header and footer](#set-header-and-footer)

```csharp
// Append a value with the given style type
odtDocument.AppendLine(long.MinValue, TextStyle.HeadingLevel01);
odtDocument.AppendLine(byte.MaxValue, TextStyle.HeadingLevel02, "Liberation Sans", 0.0);
odtDocument.AppendLine(uint.MaxValue, TextStyle.HeadingLevel03, Color.Red, Color.LightGray);
odtDocument.AppendLine(double.NaN,    TextStyle.HeadingLevel04, "Liberation Sans", 0.0, Color.Red, Color.LightGray);

// Append a value with the given style type
odtDocument.AppendLine("This is a title",  TextStyle.Title);
odtDocument.AppendLine("This is a title",  TextStyle.Title, "Liberation Sans", 0.0);
odtDocument.AppendLine("This is a title",  TextStyle.Title, Color.Red, Color.LightGray);
odtDocument.AppendLine("This is a title",  TextStyle.Title, "Liberation Sans", 0.0, Color.Red, Color.LightGray);

// Append a content with the given style type
var content = new StringBuilder();
content.Append("This is a very very very long footnote");
odtDocument.AppendLine(content, TextStyle.Footnote);
odtDocument.AppendLine(content, TextStyle.Footnote, "Liberation Sans", 0.0);
odtDocument.AppendLine(content, TextStyle.Footnote, Color.Red, Color.LightGray);
odtDocument.AppendLine(content, TextStyle.Footnote, "Liberation Sans", 0.0, Color.Red, Color.LightGray);
```

Note: Style-sheets don't support font size changes, so font size argument will be ignored

[All supported styles can find here](./styles.md)

## Append a page break

```csharp
/// Append a page break
odtDocument.AppendLines(string.Empty, TextStyle.PageBreak);

/// Append a page break before a bold text passages
odtDocument.AppendLine("This is the last sentence on page one", TextStyle.PageBreak | TextStyle.Bold);
odtDocument.AppendLine("This is the last sentence on page one", TextStyle.PageBreak | TextStyle.Bold, "Liberation Sans", 0.0);
odtDocument.AppendLine("This is the last sentence on page one", TextStyle.PageBreak | TextStyle.Bold, Color.Red, Color.LightGray);
odtDocument.AppendLine("This is the last sentence on page one", TextStyle.PageBreak | TextStyle.Bold, "Liberation Sans", 0.0, Color.Red, Color.LightGray);
```

## Append empty lines

```csharp
/// Append the given count of empty lines
odtDocument.AppendEmptyLines(10);
```

## Append a table

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

Note: Each table cell use automatic a cell style based on the given data type of the cell value

[All automatic used cell styles can be found here](./cellstyle.md)

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
