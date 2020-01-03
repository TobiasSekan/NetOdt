# NetOdt
Easy API to create ODT documents in .NET

## Simple API usage
```csharp
// Create document (C# 8.0 syntax)
using var odtDocument = new OdtDocument("E:/MyDocument.odt");

// Append a title
odtDocument.Append("My document", HeaderStyle.Title);

// Append a bold text
odtDocument.Append("This is a text", TextStyle.Bold);

// Append ten empty lines
odtDocument.AppendEmptyLines(10);

// Append a 3x3 table and fill all cells with "Fill me"
odtDocument.AppendTable(3, 3, "Fill me");

// Append a image with a width of 22.5 cm and a height of 14.1 cm
odtDocument.AppendImage("E:/picture1.jpg", 22.5, 14.1);

// The automatic dispose call (from the using syntax) do the rest of the work
// (save document, delete temporary folder, free all used resources)

// thats it :-)
```
[All API calls can find here](./api.md)

## Upcoming features
* Simple text align support (left, center, right)
* Text color support (foreground, background)
* Data type support for table cells
* ...

## Bugs report and feature requests
[Please use the GitHub issues tracker](https://github.com/TobiasSekan/NetOdt/issues)

## Supported styles
[All supported styles can find here](./styles.md)

## Goals
* Simple to use API
* Full document API and source-code
* Accept always primitive data-types, `string` and `StringBuilder` to write a value
* Use `enum` when only specific values are except
* Use `EnumFlag` for combinations of `enum` values (bold, italic, ...)
* Use build-in `DataTable` for tables
* No new data-types, use only build-in data-types
* .Net Standard 2.0, C# 8.0
