# NetOdt
Easy API to create ODT documents in .NET

## Simple API usage (short overview)
```csharp
// Needed imports
using NetOdt;
using NetOdt.Enumerations;

// Create document (C# 8.0 syntax)
using var odtDocument = new OdtDocument(path: "E:/MyDocument.odt");

// Append a title
odtDocument.Append("My document", HeaderStyle.Title);

// Append a centered bold text
odtDocument.Append("This is a bold text", TextStyle.Center | TextStyle.Bold);

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

[All supported styles can find here](./styles.md)
[All API calls can find here](./api.md)

## Upcoming features
* Header and footer
* Combinations of text style and text alignment
* Text color support (foreground, background)
* Data type support for table cells (currency for `decimal`, ...)
* ...

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
