# xlsx-toolkit

A powerful and flexible TypeScript library for loading, parsing, and modifying XLSX (Excel) files. Ideal for developers who need to manipulate Excel files in Node.js or the browser.

## Features

- Load XLSX files from local or remote sources
- Parse XLSX files into readable data structures
- Modify and update existing Excel files
- Create new XLSX files from scratch
- Export modified files for download or storage

## Installation

You can install `xlsx-toolkit` via npm:

```bash
npm install xlsx-toolkit
```
Or with yarn:
```yarn
yarn add xlsx-toolkit
```

## Basic Usage

Here's a simple example to get started:
```typescript
import { loadXLSX, parseSheet, modifySheet, saveXLSX } from 'xlsx-toolkit';

// Load an XLSX file
const workbook = loadXLSX('./example.xlsx');

// Parse the first sheet
const sheetData = parseSheet(workbook, 0);
console.log(sheetData);

// Modify the data (example: updating a cell)
sheetData[1][2] = 'Updated Value';

// Save the modified workbook
saveXLSX(workbook, './modified.xlsx');
```

## API Documentation
`loadXLSX(path: string): Workbook`
Loads an XLSX file from the specified path.

`parseSheet(workbook: Workbook, sheetIndex: number): any[]`
Parses a sheet from the workbook and returns an array of data.

`modifySheet(workbook: Workbook, sheetIndex: number, data: any[]): void`
Modifies a specific sheet within the workbook with new data.

`saveXLSX(workbook: Workbook, path: string): void`
Saves the modified workbook to the specified path.

## Contributing

Contributions are welcome! If you'd like to contribute to the development of xlsx-toolkit, feel free to open an issue or submit a pull request on GitHub.

## License

This project is licensed under the MIT License. See the LICENSE file for more details.
