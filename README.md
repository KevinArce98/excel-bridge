# @excel-bridge/core

A lightweight, modular Excel manipulation library built with micro-packages and TypeScript. Designed for performance and minimal dependencies.

## Features

- 🚀 **Lightweight** - Built with fflate and fast-xml-parser for optimal performance
- 📦 **Modular** - Micro-package architecture with tree-shaking support
- 🔒 **TypeScript** - Full type safety and IntelliSense support
- 📖 **Easy API** - Simple interface for reading and writing Excel files
- 🎯 **Browser & Node** - Works in both environments
- ⚡ **Fast** - Optimized for large datasets with shared strings support

## Installation

```bash
npm install @excel-bridge/core
```

## Quick Start

### Reading Excel Files

```typescript
import { ExcelBridge } from '@excel-bridge/core';

// From File (Browser)
const file = document.querySelector('input[type="file"]').files[0];
const workbook = await ExcelBridge.readFromFile(file);

// From Buffer (Node)
import fs from 'fs';
const buffer = fs.readFileSync('data.xlsx');
const workbook = ExcelBridge.read(buffer);

console.log(workbook.sheets[0].data);
```

### Writing Excel Files

```typescript
import { ExcelBridge } from '@excel-bridge/core';

const data = [
  ['Name', 'Age', 'City'],
  ['John', 25, 'New York'],
  ['Jane', 30, 'Los Angeles']
];

// Create Excel Blob (Browser)
const blob = ExcelBridge.write(data);
const url = URL.createObjectURL(blob);

// Create Excel Buffer (Node)
const buffer = ExcelBridge.writeBuffer(data);
import fs from 'fs';
fs.writeFileSync('output.xlsx', buffer);
```

## Advanced Usage

### Custom Styling

```typescript
import { ExcelWriter } from '@excel-bridge/core';

const writer = new ExcelWriter({ sheetName: 'Report' });

const data = [
  ['Header 1', 'Header 2'],
  ['Data 1', 'Data 2']
];

// Add styling to headers
const styledData = writer.addStyle([{ data }], 0, 0, {
  background: '#FFE0B0',
  bold: true,
  border: true
});

const blob = writer.createWorkbook(styledData);
```

### Data Validation

```typescript
import { ExcelWriter } from '@excel-bridge/core';

const writer = new ExcelWriter();

const data = [
  ['Name', 'Category'],
  ['Item 1', ''],
  ['Item 2', '']
];

// Add dropdown validation
const validatedData = writer.addValidation(
  [{ data }], 
  'B2:B3', 
  'Option1,Option2,Option3'
);

const blob = writer.createWorkbook(validatedData);
```

### Working with Coordinates

```typescript
import { coordinateToIndex, indexToCoordinate } from '@excel-bridge/core';

// Convert "A1" to indices
const { row, col } = coordinateToIndex('A1'); // { row: 0, col: 0 }

// Convert indices to "A1"
const coordinate = indexToCoordinate(0, 0); // "A1"
```

## API Reference

### Classes

- **ExcelReader** - Advanced Excel file parsing
- **ExcelWriter** - Advanced Excel file creation

### Functions

- **ExcelBridge.read(buffer)** - Parse Excel from buffer
- **ExcelBridge.readFromFile(file)** - Parse Excel from File object
- **ExcelBridge.write(data)** - Create Excel Blob from data
- **ExcelBridge.writeBuffer(data)** - Create Excel Buffer from data
- **coordinateToIndex(coord)** - Convert Excel coordinate to indices
- **indexToCoordinate(row, col)** - Convert indices to Excel coordinate

### Types

```typescript
interface ParsedCell {
  value: any;
  type: 'string' | 'number' | 'boolean' | 'empty';
  coordinate: string;
  rowIndex: number;
  columnIndex: number;
}

interface CellValidation {
  range: string;
  options: string;
}

interface CellStyle {
  background?: string;
  border?: boolean;
  bold?: boolean;
  color?: string;
}
```

## Building

```bash
npm run build    # Build for production
npm run dev      # Watch mode for development
npm run test     # Run tests
```

## Architecture

This library follows a micro-package architecture:

- **core/** - Core utilities (ZIP, XML templates, constants)
- **reader/** - Excel parsing functionality
- **writer/** - Excel creation functionality
- **index.ts** - Main entry point and unified API

## Dependencies

- **fflate** - Fast ZIP compression/decompression
- **fast-xml-parser** - High-performance XML parsing
- **TypeScript** - Type safety and development

## License

MIT
