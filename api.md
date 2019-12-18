# Full API description

_More API functions will come in next time, stay tuned_

## Important Notes

* Don't forget to use `using` syntax or call `Dispose()` when finished.
* Use `StringBuilder` for long text passages, when possible
* Use `DataTable` for tables, when possible

## Create a document

```csharp
// Create a new ODT document, save the ODT document as "Unknown.odt"
// and use a automatic generated temporary folder
using var odtDocument = new OdtDocument();

// Create a new ODT document, save the ODT document into the given file path
// and use a automatic generated temporary folder
using var odtDocument = new OdtDocument("E:/MyDocument.odt");

// Create a new ODT document, save the ODT document into the given file path
// and use the given temporary folder
using var odtDocument = new OdtDocument("E:/MyDocument.odt", "E:/TempFolder");
```

## Write unformatted values and text into the document

```csharp
// Write a single line with a unformatted value to the document
odtDocument.Write(long.MinValue);
odtDocument.Write(byte.MaxValue);
odtDocument.Write(uint.MaxValue);
odtDocument.Write(double.NaN);

// Write a single line with a unformatted text to the document
// (Note: line breaks "\n" will currently not working)
odtDocument.Write("This is a text");

// Write the content of the given StringBuilder into the document
// (Note: line breaks "\n" will currently not working)
var stringBuilder = new StringBuilder();
stringBuilder.Append("This is a text a very very very long text");
odtDocument.Write(stringBuilder);
```

## Write formatted values and text into the document

```csharp
// Write a single line with a styled value to the document
odtDocument.Write(long.MinValue, TextStyle.Bold);
odtDocument.Write(byte.MaxValue, TextStyle.Italic);
odtDocument.Write(uint.MaxValue, TextStyle.Underline);
odtDocument.Write(double.NaN, TextStyle.Bold | TextStyle.Italic | TextStyle.Underline);

// Write a single line with a unformatted text to the document
// (Note: line breaks "\n" will currently not working)
odtDocument.Write("This is a text", TextStyle.Bold | TextStyle.Underline);

// Write the content of the given StringBuilder into the document
// (Note: line breaks "\n" will currently not working)
var stringBuilder = new StringBuilder();
stringBuilder.Append("This is a text a very very very long text");
odtDocument.Write(stringBuilderFormated, TextStyle.Italic | TextStyle.Underline);
```

## Write a unformatted table

```csharp

// Write an empty unformatted table with the given row and cell count into the document
odtDocument.WriteTable(3, 3);

//  Write a unformatted table and fill it with the given data from the DataTable
var table = new DataTable();
odtDoucment.WriteTable(dataTable);
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

// Save the change content and create the ODT document into the given path and automatic override a existing file
odtDocument.SaveAs("E:/MyDocument.odt", overrideExistingFile: false);
```

## Manually call Dispose method

```csharp
// Save the document (override when existing), delete the temporary working folder and free all resources
odtDocument.Dispose();

// Save the document, delete the temporary working folder and free all resources
odtDocument.Dispose(overrideExistingFile: false);
```
