# NetOdt
Easy API to create ODT documents in .NET

## Simple API usage (short overview)
```csharp
// Needed imports
using NetOdt;
using NetOdt.Enumerations;

// Create document (C# 8.0 syntax)
using var odtDocument = new OdtDocument(path: "E:/MyDocument.odt");

// Set font (name) and font size for the complete document
odtDocument.SetGlobalFont("Liberation Serif", FontSize.Size12);

// Set header and footer
odtDocument.SetHeader("Extend documentation", TextStyle.Center | TextStyle.Bold);
odtDocument.SetFooter("Page 2/5", TextStyle.Right | TextStyle.Italic);

// Append a title
odtDocument.AppendLine("My document", TextStyle.Title);

// Append a centered bold text
odtDocument.AppendLine("This is a bold text", TextStyle.Center | TextStyle.Bold);

// Append ten empty lines
odtDocument.AppendEmptyLines(lines: 10);

// Append a 3x3 table and fill all cells with "Fill me"
odtDocument.AppendTable(rows: 3, columns: 3, content: "Fill me");

// Append a image with a width of 22.5 cm and a height of 14.1 cm
odtDocument.AppendImage(path: "E:/picture1.jpg", width: 22.5, height: 14.1);

// The automatic dispose call (from the using syntax) do the rest of the work
// (save document, delete temporary folder, free all used resources)

// thats it :-)
```

## Documentation

[All supported styles can find here](./Documentation/styles.md)

[All API calls can find here](./Documentation/api.md)

## Upcoming features
Next time, when possible:
* Text color support (foreground, background)
* Additional standard style-sheets
  * Sender
  * Receiver
  * Salutation
  * Pre-formated text
  * Horizontal line
* Data type support for table cells (currency for `decimal`, ...)
* Enumeration with typical font names

## Bugs report and feature requests
[Please use the GitHub issues tracker](https://github.com/TobiasSekan/NetOdt/issues)

## Goals
* Simple to use API
* Full document API and source-code
* Accept always primitive data-types, `string` and `StringBuilder` to write a value
* Use `enum` when only specific values are except
* Use `EnumFlag` for combinations of `enum` values (bold, italic, ...)
* Use build-in `DataTable` for tables
* No new data-types, use only build-in data-types
* .Net Standard 2.0, C# 8.0
