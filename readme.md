# NetOdt
Easy API to create ODT documents in .NET

## Simple API usage
```csharp
// Create document
using var odtDocument = new OdtDocument("E:/MyDocument.odt");

// Write a bold text
odtDocument.Write("This is a text", TextStyle.Bold);

// Write ten empty lines
odtDocument.WriteEmptyLines(10);

// Write a 3x3 table and fill all cells with "Fill me"
odtDocument.WriteTable(3, 3, "Fill me");

// The automatic dispose call (from the using syntax) do the rest of the work
// (save document, delete temporary folder, free all used resources)

// thats it :-)
```

[All API calls can find here](./api.md)

## Upcoming features
* Headline support (T1, H1 ... H10)
* Simple text align support (left, center, right)
* Text color support (foreground, background)
* Data type support for table cells
* ...

## Bugs report and feature requests
Please use the GitHub issues tracker

## Supported text styling
For style combinations you must combine the text styling via a bitwise or `|`, like
```
TextStyle.Bold | TextStyle.Italic | TextStyle.Underline
```

* Bold
* Italic
* Underline (simple)

## Goals
* Simple to use API
* Full document API and source-code
* Accept always primitive data-types, `string` and `StringBuilder` to write a value
* Use `enum` when only specific values are except
* Use `EnumFlag` for combinations of `enum` values (bold, italic, ...)
* Use build-in `DataTable` for tables
* No new data-types, use only build-in data-types
* .Net Standard 2.0, C# 8.0
