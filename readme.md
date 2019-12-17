# NetCoreOdt
Easy API to create ODT files in .NET Core applications

## Goals
* Simple to use API
* Full documentet API and source-code
* Accept always primitive datatypes, `string` and `StringBuilder` to write a value
* Use `enum` when only specific values are except
* Use `EnumFlag` for combinations of `enum` values (bold, italic, ...)
* Use build-in `DataSet` for tables
* No new datatypes, use only build-in datatypes

## API usage
```csharp
// Create and prepare a new ODT document
using var odtDoucment = new OdtDocument();

// more will come in next time, stay tuned

// Direct raw content manipulation (will be removed when features working)
odtDoucment.RawFileContent

// Direct XML content manipulation (will be removed when features working)
odtDoucment.ContentFile

// Create and prepare a new ODT document
odtDoucment.Save("E:\\testTest.odt");
```