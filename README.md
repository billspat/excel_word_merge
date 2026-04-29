
# Excel-Word SMASH!

An Electron app for someone who carefully constructs and Excel file
and an MS-Word template with placeholders

## Development

### Requirements

Node.js 18+

### Run locally
```
npm install
npm start
```
### macOS packaging
```
npm run build:mac
```

Produces `dist/TA Review Generator.app`

Unsigned apps must be opened via right-click → Open the first time.

## Using

*This should go into a help document in the app*

### Word template rules

- Use normal `.docx` and any styling you want
- Placeholders use single braces: `{ColumnName}` 

### Excel file

- tabular, first row has column names
- Column names in the excel file should match things in the template

See the help instructions in the app.  



