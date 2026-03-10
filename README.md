# @excel-bridge/core

A lightweight, modular Excel manipulation library built with micro-packages and TypeScript. Designed for performance and minimal dependencies.

## Features

- 🚀 **Lightweight** - Built with fflate and fast-xml-parser for optimal performance
- 📦 **Modular** - Micro-package architecture with tree-shaking support
- 🔒 **TypeScript** - Full type safety and IntelliSense support
- 📖 **Easy API** - Simple interface for reading and writing Excel files
- 🎯 **Browser & Node** - Works in both environments
- ⚡ **Fast** - Optimized for large datasets with shared strings support
- 🎨 **Styling Support** - Custom cell backgrounds, borders, and formatting
- ✅ **Data Validation** - Dropdown lists and validation rules
- 🔧 **Zero Heavy Dependencies** - No ExcelJS or SheetJS required

## Installation

```bash
npm install @excel-bridge/core
# or
yarn add @excel-bridge/core
# or
pnpm add @excel-bridge/core
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

## Development

### Building

```bash
npm run build         # Build for production
npm run dev           # Watch mode for development
npm run test          # Run tests with Vitest
```

### Code Quality

```bash
npm run lint          # Check code quality with ESLint
npm run lint:fix      # Auto-fix linting issues
npm run format        # Format code with Prettier
npm run format:check  # Check formatting without changes
```

### Project Structure

```
excel-bridge/
├── src/
│   ├── core/
│   │   ├── constants.ts       # XML namespaces & Office Open XML constants
│   │   ├── xml-templates.ts   # XML generators for sheets, styles, strings
│   │   └── zip-manager.ts     # ZIP compression/decompression with fflate
│   ├── reader/
│   │   └── index.ts           # Excel parsing and reading functionality
│   ├── writer/
│   │   └── index.ts           # Excel creation and writing functionality
│   └── index.ts               # Main entry point and unified API
├── tests/
│   └── basic.test.ts          # Unit tests
├── package.json
├── tsconfig.json
├── eslint.config.js           # ESLint configuration (flat config)
└── .prettierrc                # Prettier configuration
```

## Architecture

This library follows a **micro-package architecture** for optimal tree-shaking and modularity:

### Core Module (`src/core/`)
- **constants.ts** - Office Open XML namespaces and content types
- **xml-templates.ts** - XML string generators for Excel file structure
- **zip-manager.ts** - ZIP operations using fflate (compression/decompression)

### Reader Module (`src/reader/`)
- **ExcelReader class** - Parses Excel files from buffers or File objects
- **XML to JSON conversion** - Extracts cell data, validations, and metadata
- **Type-safe parsing** - Returns strongly-typed workbook structures

### Writer Module (`src/writer/`)
- **ExcelWriter class** - Creates Excel files with data, styles, and validations
- **Shared strings optimization** - Reduces file size for text-heavy data
- **Flexible API** - Simple or advanced usage patterns

### Main Entry (`src/index.ts`)
- **Unified ExcelBridge API** - Simple interface for common operations
- **Coordinate utilities** - Convert between "A1" format and array indices
- **Re-exports** - All classes, functions, and types

## Technical Details

### Dependencies

**Production:**
- **fflate** (^0.8.2) - Fast ZIP compression/decompression
- **fast-xml-parser** (^5.5.1) - High-performance XML parsing

**Development:**
- **TypeScript** (^5.9.3) - Type safety and development
- **tsup** (^8.5.1) - Ultra-fast bundler for ESM/CJS output
- **Vitest** (^4.0.18) - Modern testing framework
- **ESLint** (^10.0.3) - Code quality and linting
- **Prettier** (^3.8.1) - Code formatting

### Browser Compatibility

- Modern browsers with ES2022 support
- File API support for reading files
- Blob API support for creating downloads

### Node.js Compatibility

- Node.js ^20.19.0, ^22.13.0, or >=24
- Built with SSL support (standard in official distributions)

## Contributing

Contributions are welcome! Please ensure:

1. Code passes ESLint checks: `npm run lint`
2. Code is formatted with Prettier: `npm run format`
3. Tests pass: `npm run test`
4. TypeScript compiles without errors: `npm run build`

## License

MIT

## Author

Built with love for developers who need lightweight Excel manipulation without heavy dependencies.
