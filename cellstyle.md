# Automatic used cell styles

The open document format use a `office-value-type` to define the value type

## C# value types

| C# type       | Office value type          |
| ------------- | -------------------------- |
| bool          | `OfficeValueType.Boolean`  |
| byte          | `OfficeValueType.Float`    |
| sbyte         | `OfficeValueType.Float`    |
| short         | `OfficeValueType.Float`    |
| ushort        | `OfficeValueType.Float`    |
| int           | `OfficeValueType.Float`    |
| uint          | `OfficeValueType.Float`    |
| long          | `OfficeValueType.Float`    |
| ulong         | `OfficeValueType.Float`    |
| float         | `OfficeValueType.Float`    |
| double        | `OfficeValueType.Float`    |
| decimal       | `OfficeValueType.Currency` |
| char          | `OfficeValueType.String`   |
| string        | `OfficeValueType.String`   |

## C# object types

| C# type       | Office value type          |
| ------------- | -------------------------- |
| StringBuilder | `OfficeValueType.String`   |
| DateTime      | `OfficeValueType.Date`     |
| TimeSpan      | `OfficeValueType.Time`     |
