# Automatic used cell styles

The open document format (*.ODT) use a `office:value-type` to define the value type

Note: The same `office-value-type`s can have different styles for number presentations

## C# value types

| C# type       | NetOdt type                | office:value-type | ODT value name       | Additional attributes |
| ------------- | -------------------------- | ----------------- | -------------------- | --------------------- |
| bool          | `OfficeValueType.Boolean`  | boolean           | office:boolean-value |                       |
| byte          | `OfficeValueType.Float`    | float             | office:value         |                       |
| sbyte         | `OfficeValueType.Float`    | float             | office:value         |                       |
| short         | `OfficeValueType.Float`    | float             | office:value         |                       |
| ushort        | `OfficeValueType.Float`    | float             | office:value         |                       |
| int           | `OfficeValueType.Float`    | float             | office:value         |                       |
| uint          | `OfficeValueType.Float`    | float             | office:value         |                       |
| long          | `OfficeValueType.Float`    | float             | office:value         |                       |
| ulong         | `OfficeValueType.Float`    | float             | office:value         |                       |
| float         | `OfficeValueType.Float`    | float             | office:value         |                       |
| double        | `OfficeValueType.Float`    | float             | office:value         |                       |
| decimal       | `OfficeValueType.Currency` | currency          | office:value         | office:currency="EUR" |
| char          | `OfficeValueType.String`   | string            | (not need)           |                       |
| string        | `OfficeValueType.String`   | string            | (not need)           |                       |

## C# object types

| C# type       | NetOdt type                | office:value-type | ODT value name       | ODT value format    |
| ------------- | -------------------------- | ----------------- | -------------------- | ------------------- |
| StringBuilder | `OfficeValueType.String`   | string            | (not need)           |                     |
| DateTime      | `OfficeValueType.Date`     | date              | office:date-value    | `{yyyy}-{MM}-{dd}`  |
| TimeSpan      | `OfficeValueType.Time`     | time              | office:time-value    | `PT{hh}H{mm}M{ss}S` |

## None C# types
| Office value type            | ODT type     | ODT value name       |
| ---------------------------- | ------------ | -------------------- |
| `OfficeValueType.Percentage` | percentage   | office:value         |
| `OfficeValueType.Scientific` | float        | office:value         |
| `OfficeValueType.Fraction`   | fraction     | office:value         |
