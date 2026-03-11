# excel-bridge

A powerful, lightweight Excel manipulation library built with TypeScript and micro-packages. Features advanced styling, multi-sheet support, formulas, and more - all with zero heavy dependencies.

## ✨ Advanced Features

- 🚀 **High Performance** - Built with fflate and fast-xml-parser for optimal speed
- 📦 **Modular Architecture** - Tree-shaking support with micro-packages
- 🔒 **Full TypeScript** - Complete type safety and IntelliSense
- 📖 **Simple API** - Intuitive interface for reading and writing Excel files
- 🎯 **Cross-Platform** - Works in both Browser and Node.js environments
- 🎨 **Dynamic Styling** - Advanced StyleManager with colors, fonts, borders
- 📊 **Multi-Sheet Support** - Create workbooks with multiple named sheets
- 🧮 **Formula Support** - Native Excel formulas (=SUM, =TODAY, custom formulas)
- 📅 **Date Handling** - Automatic conversion to Excel serial dates
- 🔄 **Merge Cells** - Combine cells across ranges
- 🧊 **Freeze Panes** - Lock header rows/columns
- 📏 **Auto-Width Columns** - Automatic column width calculation
- ✅ **Data Validation** - Dropdown lists and validation rules
- 🔧 **Zero Heavy Dependencies** - No ExcelJS or SheetJS required

## Installation

```bash
npm install excel-bridge
# or
yarn add excel-bridge
# or
pnpm add excel-bridge
```

## Quick Start

### Reading Excel Files

```typescript
import { ExcelBridge } from 'excel-bridge';

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
import { ExcelBridge } from 'excel-bridge';

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

### Advanced Multi-Sheet Example

```typescript
import { ExcelWriter } from 'excel-bridge';

const writer = new ExcelWriter({ creator: 'My App' });

// Sheet 1: Sales data with styling and formulas
const salesSheet = {
  data: [
    ['Product', 'Price', 'Quantity', 'Total'],
    ['Laptop', 999.99, 5, '=B2*C2'],
    ['Mouse', 29.99, 20, '=B3*C3'],
    ['Keyboard', 79.99, 10, '=B4*C4']
  ],
  styles: {
    '0-0': { background: '#4472C4', bold: true, color: '#FFFFFF' },
    '0-1': { background: '#4472C4', bold: true, color: '#FFFFFF' },
    '0-2': { background: '#4472C4', bold: true, color: '#FFFFFF' },
    '0-3': { background: '#4472C4', bold: true, color: '#FFFFFF' },
  },
  options: {
    name: 'Sales Report',
    freezePane: { row: 1 },
    autoWidth: true
  }
};

// Sheet 2: Dates and calculations
const datesSheet = {
  data: [
    ['Event', 'Date', 'Days Until Today'],
    ['Launch', new Date(2024, 6, 15), '=TODAY()-B2'],
    ['Meeting', new Date(2024, 8, 20), '=TODAY()-B3'],
    ['Deadline', new Date(2024, 11, 31), '=TODAY()-B4']
  ],
  options: {
    name: 'Timeline',
    autoWidth: true
  }
};

// Create workbook with multiple sheets
const buffer = writer.createWorkbookBuffer([salesSheet, datesSheet]);
import fs from 'fs';
fs.writeFileSync('advanced-report.xlsx', buffer);
```

## Advanced Usage

### Dynamic Styling with StyleManager

```typescript
import { ExcelWriter, StyleManager } from 'excel-bridge';

const writer = new ExcelWriter({ creator: 'My App' });

const sheet = {
  data: [
    ['Product', 'Price', 'Stock', 'Status'],
    ['Laptop', 999.99, 15, 'Available'],
    ['Mouse', 29.99, 5, 'Low Stock'],
    ['Keyboard', 79.99, 0, 'Out of Stock']
  ],
  styles: {
    // Header styling
    '0-0': { background: '#4472C4', bold: true, color: '#FFFFFF' },
    '0-1': { background: '#4472C4', bold: true, color: '#FFFFFF' },
    '0-2': { background: '#4472C4', bold: true, color: '#FFFFFF' },
    '0-3': { background: '#4472C4', bold: true, color: '#FFFFFF' },
    
    // Conditional styling
    '2-2': { background: '#FFC7CE', color: '#9C0006' }, // Low stock - red
    '3-3': { background: '#FFE6E6', color: '#C00000' }, // Out of stock
    '1-2': { background: '#E2EFDA', color: '#006100' },  // Good stock - green
  },
  options: {
    name: 'Inventory',
    freezePane: { row: 1 },
    autoWidth: true
  }
};

const buffer = writer.createWorkbookBuffer([sheet]);
```

### Merge Cells & Advanced Layout

```typescript
import { ExcelWriter } from 'excel-bridge';

const writer = new ExcelWriter();

const reportSheet = {
  data: [
    ['Q1 2024 Sales Report', '', '', ''],
    ['', '', '', ''],
    ['Product', 'January', 'February', 'March'],
    ['Laptops', 45000, 52000, 48000],
    ['Accessories', 12000, 15000, 13500],
    ['Software', 8000, 8500, 9000],
    ['', '', '', ''],
    ['TOTAL', '=SUM(B4:B6)', '=SUM(C4:C6)', '=SUM(D4:D6)']
  ],
  styles: {
    '0-0': { background: '#5B9BD5', bold: true, color: '#FFFFFF' },
    '7-0': { background: '#70AD47', bold: true, color: '#FFFFFF' },
    '7-1': { background: '#70AD47', bold: true, color: '#FFFFFF' },
    '7-2': { background: '#70AD47', bold: true, color: '#FFFFFF' },
    '7-3': { background: '#70AD47', bold: true, color: '#FFFFFF' },
  },
  mergeCells: ['A1:D1', 'A8:D8'], // Merge title and total rows
  options: {
    name: 'Quarterly Report',
    freezePane: { row: 3 },
    autoWidth: true
  }
};

const buffer = writer.createWorkbookBuffer([reportSheet]);
```

### Date Handling & Excel Formulas

```typescript
import { ExcelWriter } from 'excel-bridge';

const writer = new ExcelWriter();

const projectSheet = {
  data: [
    ['Task', 'Start Date', 'End Date', 'Duration', 'Status'],
    ['Design', new Date(2024, 0, 15), new Date(2024, 1, 20), '=C2-B2', 'Completed'],
    ['Development', new Date(2024, 1, 21), new Date(2024, 4, 30), '=C3-B3', 'In Progress'],
    ['Testing', new Date(2024, 5, 1), new Date(2024, 5, 15), '=C4-B4', 'Planned'],
    ['Deployment', new Date(2024, 5, 16), new Date(2024, 5, 20), '=C5-B5', 'Planned'],
    ['', '', '', '', ''],
    ['Project Duration', '', '', '=MAX(D2:D5)', ''],
    ['Tasks Completed', '', '', '=COUNTIF(E2:E5,"Completed")', '']
  ],
  options: {
    name: 'Project Timeline',
    freezePane: { row: 1 },
    autoWidth: true
  }
};

const buffer = writer.createWorkbookBuffer([projectSheet]);
```

### Working with Coordinates

```typescript
import { coordinateToIndex, indexToCoordinate } from 'excel-bridge';

// Convert "A1" to indices
const { row, col } = coordinateToIndex('A1'); // { row: 0, col: 0 }

// Convert indices to "A1"
const coordinate = indexToCoordinate(0, 0); // "A1"
```

## API Reference

### Classes

- **ExcelReader** - Advanced Excel file parsing with full feature support
- **ExcelWriter** - Advanced Excel file creation with StyleManager
- **StyleManager** - Dynamic style management for optimal performance

### Core Functions

- **ExcelBridge.read(buffer)** - Parse Excel from buffer
- **ExcelBridge.readFromFile(file)** - Parse Excel from File object
- **ExcelBridge.write(data)** - Create Excel Blob from data
- **ExcelBridge.writeBuffer(data)** - Create Excel Buffer from data
- **coordinateToIndex(coord)** - Convert Excel coordinate to indices
- **indexToCoordinate(row, col)** - Convert indices to Excel coordinate

### Utility Functions

- **dateToExcelSerial(date)** - Convert JavaScript Date to Excel serial number
- **calculateColumnWidths(data)** - Calculate optimal column widths
- **escapeXml(text)** - Escape XML special characters

### Types

```typescript
interface ExcelData {
  data: any[][];
  styles?: Record<string, CellStyle>;
  validations?: CellValidation[];
  mergeCells?: string[];
  options?: SheetOptions;
}

interface SheetOptions {
  name?: string;
  freezePane?: { row: number; col: number };
  autoWidth?: boolean;
}

interface CellStyle {
  background?: string;
  border?: boolean;
  bold?: boolean;
  color?: string;
  italic?: boolean;
  fontSize?: number;
  fontName?: string;
}

interface CellValidation {
  range: string;
  options: string;
}

interface ExcelWriterOptions {
  creator?: string;
  title?: string;
}

interface ParsedCell {
  value: any;
  type: 'string' | 'number' | 'boolean' | 'empty';
  coordinate: string;
  rowIndex: number;
  columnIndex: number;
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
│   │   ├── xml-templates.ts   # XML generators for sheets, styles, workbooks
│   │   ├── style-manager.ts   # Dynamic style management system
│   │   ├── date-utils.ts      # Excel serial date conversion utilities
│   │   ├── column-width.ts    # Auto-width calculation algorithms
│   │   └── zip-manager.ts     # ZIP compression/decompression with fflate
│   ├── reader/
│   │   └── index.ts           # Excel parsing and reading functionality
│   ├── writer/
│   │   └── index.ts           # Advanced Excel creation with StyleManager
│   └── index.ts               # Main entry point and unified API
├── tests/
│   ├── basic.test.ts          # Core functionality tests
│   ├── read-write.test.ts     # Read/Write integration tests
│   ├── zip-structure.test.ts  # ZIP structure validation
│   ├── numbers-only.test.ts    # Numeric data handling
│   └── special-characters.test.ts # Special character handling
├── FEATURES.md                # Detailed feature documentation
├── package.json
├── tsconfig.json
├── eslint.config.js           # ESLint configuration (flat config)
└── .prettierrc                # Prettier configuration
```

## Architecture

This library follows a **micro-package architecture** for optimal tree-shaking and modularity with advanced Excel features:

### Core Module (`src/core/`)
- **constants.ts** - Office Open XML namespaces and content types
- **xml-templates.ts** - XML generators for sheets, styles, workbooks, relationships
- **style-manager.ts** - Dynamic style management with deduplication and optimization
- **date-utils.ts** - Excel serial date conversion and validation
- **column-width.ts** - Intelligent column width calculation algorithms
- **zip-manager.ts** - ZIP operations using fflate (compression/decompression)

### Reader Module (`src/reader/`)
- **ExcelReader class** - Parses Excel files from buffers or File objects
- **XML to JSON conversion** - Extracts cell data, validations, styles, and metadata
- **Type-safe parsing** - Returns strongly-typed workbook structures
- **Multi-sheet support** - Handles complex workbook structures

### Writer Module (`src/writer/`)
- **ExcelWriter class** - Creates Excel files with advanced features
- **StyleManager integration** - Optimized style generation and management
- **Multi-sheet creation** - Support for multiple named sheets with options
- **Formula support** - Native Excel formula handling with placeholders
- **Date handling** - Automatic JavaScript Date to Excel serial conversion
- **Layout features** - Merge cells, freeze panes, auto-width columns
- **Validation** - Data validation and sanitization

### Main Entry (`src/index.ts`)
- **Unified ExcelBridge API** - Simple interface for common operations
- **Advanced exports** - All classes, utilities, and types for power users
- **Coordinate utilities** - Convert between "A1" format and array indices
- **Re-exports** - Complete API surface for tree-shaking

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

### 📋 Known Limitations
- **Standard Warnings** - Excel may show warnings when opening programmatically generated files (harmless)
- **No Shared Strings** - Uses inline strings for simplicity and reliability
- **Excel 2016+** - Requires modern Excel versions for full compatibility

### 🚀 Performance
- **Lightweight** - ~35KB minified bundle size
- **Fast Processing** - Optimized for large datasets
- **Memory Efficient** - Stream-based processing for big files
- **Tree Shakable** - Import only what you need

## License

MIT

## Author

Built with ❤️ for developers who need powerful Excel manipulation without heavy dependencies.

---

**Excel Bridge** - Advanced Excel features, zero dependencies, maximum performance.
