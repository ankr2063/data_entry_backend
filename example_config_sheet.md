# Configuration Sheet Example

Create a second worksheet named "Config" with the following structure:

| Field Name | Type | Required | Validation | Options | Placeholder |
|------------|------|----------|------------|---------|-------------|
| Date | date | true | | | Select date |
| Shift | select | true | | Day,Night | Choose shift |
| Temperature (°C) | number | true | min:0;max:100 | | Enter temperature |
| Humidity (%) | number | true | min:0;max:100 | | Enter humidity % |
| Class Manufacturer | text | false | maxlength:50 | | Manufacturer name |
| Force Per Shift | select | true | | 5% < 1% C (Pre-Lamination level),8% ≤ 1% C (Lamination level) | Select force level |

## Validation Format
- `min:value` - Minimum value
- `max:value` - Maximum value  
- `minlength:value` - Minimum text length
- `maxlength:value` - Maximum text length
- `pattern:regex` - Regular expression pattern

## Field Types
- `text` - Single line text
- `textarea` - Multi-line text
- `number` - Numeric input
- `decimal` - Decimal number
- `date` - Date picker
- `email` - Email input
- `boolean` - Checkbox
- `select` - Dropdown selection