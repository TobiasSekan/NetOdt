# NetCoreOdt
Easy API to create ODT files in .NET Core applications

## Simple API usage
```csharp
// Create document
using var odtDocument = new OdtDocument("E:/MyDocument.odt");

// Write into document
odtDocument.Write("This is a text", TextStyle.Bold);

// automatic dispose call will do the rest of the work

// thats it :-)
```

[All API calls can find here](./api.md)

## Upcoming features
* Simple Text align support (left, center, right)
* Simple table support
* Text color support (foreground, background)
* ...

## Bugs report and feature requests
Please use the GitHub issues tracker

## Supported text styling
For style combinations you must combine the text styling via a bitwise or `|`,
like `TextStyle.Bold | TextStyle.Italic | TextStyle.Underline`

* Bold
* Italic
* Underline (simple)

## Goals
* Simple to use API
* Full document API and source-code
* Accept always primitive data-types, `string` and `StringBuilder` to write a value
* Use `enum` when only specific values are except
* Use `EnumFlag` for combinations of `enum` values (bold, italic, ...)
* Use build-in `DataSet` for tables
* No new data-types, use only build-in data-types
* .Net Standard 2.1, C# 8.0
