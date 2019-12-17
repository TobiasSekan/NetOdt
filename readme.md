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
// using var odtDoucment = new OdtDocument("C:/tempFolder");

// ****************************************************
// more API function will come in next time, stay tuned
// ****************************************************

// Temporary direct raw content manipulation (will be removed when features working)
odtDoucment.RawFileContent

// Temporary direct XML content manipulation (will be removed when features working)
odtDoucment.ContentFile

// Save the ODT document into the given path
odtDoucment.Save("C:/testTest.odt");

// ODT document will dispose here automatically (and remove temporary working folder)
```