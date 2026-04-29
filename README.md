
# TA Review Generator (Electron, Node-only)

## Requirements
- Node.js 18+

## Run locally
```
npm install
npm start
```

## Word template rules
- Use normal `.docx`
- Placeholders use single braces: `{ColumnName}`
- Column names must match Excel headers exactly

## macOS packaging
```
npm run build:mac
```

Produces `dist/TA Review Generator.app`

Unsigned apps must be opened via right-click → Open the first time.
