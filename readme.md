# NetCoreOdt
Easy API to create ODT files in .NET Core applications

## Goals
* Simple to use API
* Full document API and source-code
* Accept always primitive data-types, `string` and `StringBuilder` to write a value
* Use `enum` when only specific values are except
* Use `EnumFlag` for combinations of `enum` values (bold, italic, ...)
* Use build-in `DataSet` for tables
* No new data-types, use only build-in data-types
* .Net Standard 2.1, C# 8.0

## API usage (work in progress)
```csharp
// Create and prepare a new ODT document
using var odtDoucment = new OdtDocument();

// Create and prepare a new ODT document inside the given temporary folder
// using var odtDoucment = new OdtDocument("E:\\testTest.odt");

// Write a single line with a unformatted value to the document
odtDoucment.Write(long.MinValue);
odtDoucment.Write(byte.MaxValue);
odtDoucment.Write(uint.MaxValue);
odtDoucment.Write(double.NaN);

// Write a single line with a unformatted text to the document (Note: line breaks "\n" will currently not working)
odtDoucment.Write("This is a text");

// Write the content of the given <see cref="StringBuilder"/> into the document (Note: line breaks "\n" will currently not working)
var stringBuilder = new StringBuilder();
stringBuilder.Append("This is a text a very very very long text");
odtDoucment.Write(stringBuilder);

// Save the ODT document into the given path
odtDoucment.Save("E:\\testTest.odt");

// ODT document will dispose here automatically (and remove temporary working folder)

// *****************************************************
// more API functions will come in next time, stay tuned
// *****************************************************
```