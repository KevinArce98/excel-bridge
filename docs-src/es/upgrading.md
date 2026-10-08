---
title: Actualización
description: Qué cambia entre las versiones de excel-bridge y qué hacer al respecto, de 1.x a 2.0 y en versiones anteriores.
group: Referencia
groupOrder: 3
order: 4
---

# Actualización

Notas sobre los cambios que pueden afectar al código escrito para una versión anterior. Las notas de la versión enumeran qué cambió; esta página dice qué hacer al respecto.

## De 1.x a 2.0

La versión 2.0 cambia lo que significa una cadena simple, lo que devuelven el lector y `Workbook`, y lo que exporta el paquete. La mayoría del código necesita una búsqueda y unas pocas ediciones. Haz el primer paso en 1.6, antes de actualizar.

### Antes de actualizar: escribe las fórmulas como objetos

En 1.x, una cadena que empieza con `=` se escribía como fórmula. En 2.0 es texto. El cambio falla en silencio: una celda que tenía `=SUM(A1:A3)` como fórmula ahora muestra esos caracteres como texto.

| Escribes | 1.x | 2.0 |
| --- | --- | --- |
| `'=SUM(A1:A3)'` | una fórmula | el texto `=SUM(A1:A3)` |
| `{ formula: 'SUM(A1:A3)' }` | una fórmula | una fórmula, los mismos bytes |
| `{ formula: 'x', result: 4 }` | una fórmula con un valor guardado | lo mismo |
| `{ text: '=SUM(A1:A3)' }` | texto | texto, los mismos bytes que la cadena simple |

`{ formula }` y `{ text }` ya existen en 1.6 y significan lo mismo allí, así que puedes migrar primero y actualizar después. Los archivos no cambian.

Encuentra las cadenas con esta regla de ESLint:

```js
{
  rules: {
    'no-restricted-syntax': [
      'error',
      {
        selector: 'Literal[value=/^=/]:not(Property[key.name="text"] > Literal)',
        message: "A string that starts with '=' is text in excel-bridge 2.0. Use { formula: '...' } for a formula.",
      },
      {
        selector: 'TemplateLiteral > TemplateElement.quasis:first-child[value.raw=/^=/]',
        message: "A string that starts with '=' is text in excel-bridge 2.0. Use { formula: '...' } for a formula.",
      },
    ],
  },
}
```

Si las fórmulas te llegan como cadenas desde un archivo de configuración, una base de datos o una API, conviértelas al recibirlas. Esto vuelve a activar la inyección de fórmulas para esos datos, así que hazlo solo con datos de confianza.

```ts
import type { CellValue } from 'excel-bridge';

const asFormulas = (rows: CellValue[][]): CellValue[][] =>
  rows.map(row =>
    row.map(cell =>
      typeof cell === 'string' && cell.startsWith('=') ? { formula: cell.slice(1) } : cell
    )
  );
```

La ventaja es que exportar datos de usuarios ya no puede crear una fórmula por accidente. Un texto como `=== Summary ===` ya no necesita un envoltorio.

### Escritura

- **`CellValidation.options` desapareció**, tanto en las reglas que escribes como en `ParsedSheet.validations`. Una regla de lista usa `formula1`. `dataValidation.list(range, values)` no cambia para quien la llama. Una lista escrita a mano necesita `formula1: '"a,b"'`, y una lista sin `formula1` lanza `Validation at A2:A5 needs formula1`. Para leer los valores de una lista en línea: `rule.formula1?.match(/^"(.*)"$/s)?.[1].replace(/""/g, '"').split(',')`.
- **Se eliminan `ExcelWriter.addValidation`, `addStyle`, `createSimple` y `createSimpleBuffer`.** Define tú mismo `data[data.length - 1].validations` y `.styles`, o usa `dataValidation`, y usa `createExcelFile` o `createExcelFileBuffer` para una sola hoja.
- **`SheetLayout` está completo.** Ahora contiene `freezePane` y `columnWidths`, además de `rowHeights`, `hiddenRows` y `hiddenColumns`. `calculateColumnWidths` recibe `CellValue[][]`.

### Lectura

- **`ParsedCell` es una unión por `type`.** `cell.value` ya no es `any`. Acota con `cell.type` y `value` tiene el tipo correcto. Una celda `'error'` tiene `ExcelErrorValue | string`. Una celda de fórmula tiene el tipo y el valor de su resultado guardado, y `formula` no existe o no está vacía.
- **`ParsedSheet.data` se indexa por la posición de la fila.** `data[4]` es la fila 5. Una fila que no está en el archivo es un hueco, `data.length` es la última fila más uno, y un elemento `<row>` vacío también es un hueco, así que un ciclo de ida y vuelta ya no agrega filas. Usa `forEach`, `Object.values` o `flat()`. Un `for...of` produce `undefined` en un hueco, usar spread en el arreglo convierte cada hueco en `undefined`, y `JSON.stringify` escribe `null` por cada uno: un archivo cuya única fila está en el índice 1,000,000 produce más de 5 millones de caracteres (5,000,090 con una celda numérica).
- **Las filas y las celdas sin atributo `r`** reciben la posición siguiente a la anterior, así que `rowIndex` nunca es `NaN`. Las filas duplicadas se fusionan, y gana la celda que aparece después.
- **Un seguidor de fórmula compartida no tiene `formula`.** Se lee como su valor guardado, que `Workbook` guarda como un valor simple.
- **`CellStyle` reemplaza a `ParsedCellStyle`.** Un estilo leído de un archivo se puede volver a escribir tal cual. `CellStyle` acepta un argumento de tipo opcional para el borde, con un valor predeterminado, así que los usos existentes siguen funcionando.

### Workbook

- **`getCellValue` y `getSheetData` devuelven lo que cargan.** Una fórmula cargada es `{ formula, result? }` y un error cargado es `{ error }`. El texto sigue siendo texto, incluido el que empieza con `=`. El código que comparaba un valor con `'#N/A'` o imprimía `String(value)` para una celda de fórmula ahora ve un objeto. `typeof value === 'object' && value !== null && 'formula' in value` acota a una fórmula, y `'error' in value` a un error.
- **Se conservan los resultados guardados de las fórmulas.** 1.6 los descartaba al guardar. Un resultado no cambia cuando editas las celdas de las que depende, así que puede quedar desactualizado hasta que Excel recalcule al abrir. Para descartar uno, asigna a la celda `{ formula: cell.formula }`.
- **`splice` y `unshift` en las filas de `getSheetData` son seguros.** Ya no se rastrea nada por posición.

### Errores y límites

- **Todo lo que la biblioteca lanza a propósito es un `ExcelBridgeError`** con un `code` (`INVALID_INPUT`, `INVALID_FILE`, `LIMIT_EXCEEDED` o `UNSUPPORTED`). Los mensajes no cambian, salvo el del límite de celdas que aparece más abajo. `error.name` ahora es `ExcelBridgeError`, así que el código que lo comparaba con `'Error'` debe usar `isExcelBridgeError(error)` o `error instanceof Error`. Un error de `parseFromBuffer` conserva el prefijo `Failed to parse Excel file:` y deja el error original en `cause`.
- **`new ExcelReader()` tiene límites.** Rechaza los archivos que superan `maxCells: 5_000_000` (ahora cuenta todas las celdas, no solo el relleno), `maxPartBytes: 268_435_456` y `maxTotalBytes: 536_870_912`, con `LIMIT_EXCEEDED`. Una parte que declara un tamaño mayor que su límite se rechaza antes de descomprimirla. Pasa `Infinity` en cada uno para quitar los topes. El mensaje `Workbook pads more than 5000000 empty cells to keep rows rectangular` ahora es `Workbook has at least N cells, counting the empty cells that pad rows, over the limit maxCells of 5000000`.
- **El lector de XML es más estricto.** Una etiqueta que no coincide o que no se cierra, un atributo sin comillas y un documento que termina antes de tiempo ahora lanzan `INVALID_FILE`, donde 1.x leía lo que podía. Una parte con `DOCTYPE` o en UTF-16 se rechaza. Una parte `docProps` mal formada se omite, así que sus metadatos quedan vacíos. Una referencia numérica de carácter como `&#233;` se decodifica.
- **Los archivos con espacios de nombres XML con prefijo se leen correctamente.** Una hoja escrita como `<x:worksheet>` ya no se lee como vacía.
- **Las celdas `t="str"` con `xml:space="preserve"`**, como las escribe SheetJS, se leen como su texto en lugar de `[object Object]`.

### Exportaciones eliminadas

| Eliminado | Usa en su lugar |
| --- | --- |
| `StyleManager` | Nada: solo servía a los generadores de XML. |
| `generateSheetXml`, `generateSharedStringsXml`, `generateStylesXml`, `generateContentTypesXml`, `generateWorkbookXml`, `generateWorkbookRelsXml`, `generateRootRelsXml`, `generateCorePropsXml`, `generateAppPropsXml`, `generateSheetRelsXml`, `generateColsXml` | `ExcelWriter` o `createExcelWorkbookStream` para archivos completos. Las partes no se pueden usar por separado. |
| `createExcelBlob`, `createExcelBuffer`, `extractExcelFiles`, `validateExcelStructure` | `fflate` directamente (`zipSync`, `unzipSync`, `strToU8`, `strFromU8`). Ya es una dependencia. |
| `XML_NS`, `CONTENT_TYPES`, `RELATIONSHIP_TYPES`, `CELL_TYPES` | Nada: la biblioteca los fija en su código. |
| `isDateNumFmtId`, `isDateFormatCode`, `validateRowIndex`, `validateColIndex`, `validateCellValue` | Nada: son detalles internos del lector y del escritor. |
| Tipos `ParsedCellStyle`, `Font`, `Fill`, `Border`, `ExcelStyle`, `CellAlignment`, `ExcelFiles`, `SheetGenerationOptions`, `DefinedName` | `CellStyle` para los estilos. Los demás no tienen reemplazo. |

`parseExcel`, `createExcelFile`, `createExcelFileBuffer` y `ExcelBridge` se mantienen. `isExcelError` es nuevo.

### Novedades

- `sheetToObjects`, `objectsToSheet` y `objectsToStreamingSheet`: filas como objetos con tipo.
- `downloadXlsx`, `toReadableStream`, `xlsxResponse` y `XLSX_CONTENT_TYPE`: enviar un archivo a un navegador o desde un servidor.
- `ExcelBridgeError`, `isExcelBridgeError` y `isExcelError`.
- Opciones del lector en `ExcelReader`, `parseExcel`, `ExcelBridge.read`, `ExcelBridge.readFromFile`, `Workbook.fromBuffer` y `Workbook.fromFile`.

### Tamaño y velocidad

Tamaño del paquete, min+gzip, medido con la misma llamada de esbuild en ambas versiones:

| Entrada | 1.6.0 | 2.0 |
| --- | ---: | ---: |
| `createExcelWorkbookStream` | 12.3 KB | 12.3 KB |
| `ExcelWriter` | 12.9 KB | 12.8 KB |
| `ExcelReader` | 28.8 KB | 9.5 KB |
| `Workbook` | 41.0 KB | 21.5 KB |
| `ExcelBridge` | 41.2 KB | 21.6 KB |
| Todo | 43.7 KB | 25.1 KB |

Leer un archivo de 50,000 × 10 tarda unos 0.27 s en lugar de 1.5 s, y el pico de memoria del proceso del benchmark baja de 859 a 451 MiB (Node 24.19, Apple M4). El lector es más rápido porque analiza el XML con su propio tokenizador, y el paquete ahora depende solo de `fflate`. La escritura es tan rápida como antes.

## De 1.5 a 1.6

Todo lo que podías escribir en 1.5 sigue escribiendo el mismo archivo, salvo por las notas de salida que aparecen más abajo. Las funciones nuevas son aditivas. Cuatro tipos de cambio pueden afectar al código: tipos de TypeScript que se amplían, unas pocas llamadas que ahora lanzan un error, valores de salida y del lector que difieren, y ciclos de ida y vuelta de `Workbook` que conservan más.

### Tipos que se amplían (solo TypeScript)

Los valores en tiempo de ejecución de estos no cambian, salvo `ParsedSheet.styles[...].border` (consulta "Valores que devuelve el lector"), pero el código que acota los tipos anteriores puede dejar de compilar.

- **`CellValue`** ahora también cubre `FormulaCell`, `TextCell` y `ErrorCell`. El código que lee valores (`getCellValue`, `getSheetData`) y comprueba `typeof value === 'object'` para encontrar un `Date` necesita `value instanceof Date`. Asignar `getSheetData()` a un arreglo de la unión anterior (`string | number | boolean | Date | null | undefined`) requiere `CellValue`.
- **Los literales de arreglo de objetos de celda**, como `[[{ error: '#N/A' }]]`, infieren `string` para el error. Anota el arreglo como `CellValue[][]`, o escribe `{ error: '#N/A' as const }`.
- **`CellStyle.border`** es `CellBorder`: `boolean`, un estilo de línea, `{ style, color? }` o entradas por lado. `ParsedSheet.styles[...].border` es `true` para el recuadro delgado simple y, en cualquier otro caso, un objeto con una entrada por lado, así que `const border: boolean | undefined = style.border` ya no compila. Las comprobaciones de valor verdadero siguen funcionando.
- **`ParsedSheet`, `SheetOptions` y `StreamingSheetInput`** incorporan `rowHeights`, `hiddenRows` y `hiddenColumns`. Una comprobación exhaustiva sobre sus claves necesita las nuevas.
- **`generateColsXml(widths, layout?)`** recibe un objeto de diseño como segundo argumento e ignora un número, así que pasarla directamente a `Array.prototype.map` sigue ejecutándose pero ya no pasa la comprobación de tipos. Una subclase de `StyleManager` que declara sus propios miembros privados `registerStyle`, `styleIds` o `internXf` entra en conflicto con los nuevos miembros privados.

### Llamadas que ahora lanzan un error

| Entrada | Qué hacer |
| --- | --- |
| `{ formula: '' }` (o `{ formula: '=' }`) | Pasa una fórmula. La abreviatura de cadena `'='` no cambia. |
| `{ error: '#SPILL!' }` o cualquier error fuera de los siete clásicos | Usa `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!` o `#N/A`, o escribe el texto con `{ text }`. |
| Una altura de fila de 0 o mayor que 409.5, o un índice de fila o columna oculta que no es un número entero (filas de 0 a 1,048,575, columnas de 0 a 16,383) | Corrige el valor. |
| Un estilo de línea de borde distinto de los 13 nombres, un color que no es hexadecimal, o un valor verdadero que no es un borde (`1`, `'true'`, `{ left: true }`) | Usa `true`, un estilo de línea, `{ style, color? }` o entradas por lado. |
| Un objeto de borde que mezcla `color` con claves de lado (`{ left: 'thin', color: '#000000' }`) | Pon el color dentro de cada lado. TypeScript también rechaza `{ left: 'thin', style: 'thin' }`; en tiempo de ejecución se ignoran sus claves de lado. |

Un objeto con claves `formula`, `text` o `error` se escribía como el texto `[object Object]`, y `{ formula: 1 }` o `{ text: 1 }` ahora lanzan `Cell A1 needs a string formula or text`. `{}` y `[]` como `border` (sin tipo, solo JavaScript) dibujaban un recuadro y ahora significan que no hay borde.

### Salida

- **Los estilos iguales comparten una entrada.** Los estilos redundantes (indicadores `false` explícitos, `#f00` junto a `#FF0000`, `fontName: 'Calibri'` escrito de forma explícita, estilos de formato condicional iguales) ahora se reducen a uno, así que los conteos de `styles.xml` y los índices `s=` pueden cambiar. `ExcelReader` devuelve los mismos valores y estilos. `border: false` (o `null`, `0`, `''`) ya no agrega un segundo borde vacío.
- **`generateStylesXml()` sin argumento** devuelve lo mismo que `generateStylesXml(new StyleManager())`, que es el `styles.xml` de un archivo de `ExcelWriter` sin estilos.
- **Se escribe el estilo de una celda `Date`.** La negrita, el relleno, el borde, la alineación y un `numberFormat` de fecha se ignoraban. Un `numberFormat` que no es un formato de fecha se sigue ignorando en una celda `Date`, y un hipervínculo sobre una celda `Date` ahora le da la fuente de hipervínculo.
- **Una columna oculta sin ancho** se escribe con ancho `9.140625`, el ancho de columna predeterminado en los archivos de Excel.
- **Tamaño:** `ExcelWriter` pasa de 12.5 a 12.9 KB min+gzip, el escritor de streaming de 11.9 a 12.3 KB, `ExcelReader` de 28.2 a 28.7 KB y `Workbook` de 39.9 a 41.0 KB.

### Valores que devuelve el lector

- **Las celdas con un formato de número de fecha informan su estilo** cuando además tiene un formato de número personalizado, fuente, relleno, borde o alineación, sea cual sea el tipo de celda (una cadena o una celda vacía pueden llevar uno), y `numberFormat` lleva el código de los ids integrados del 15 al 22 y del 45 al 47. Una celda de fecha corta simple sigue sin informar ninguno, y `Workbook.getCellStyle` cambia de la misma manera.
- **`border` en un estilo leído** era `true` para cualquier borde. Ahora es `true` solo para una línea negra delgada en los cuatro lados y, en cualquier otro caso, un objeto con una entrada `{ style, color? }` por cada lado que tiene una línea, por ejemplo `{ left: { style: 'thin' } }`. Reemplaza `style.border === true` por una comprobación de valor verdadero.
- **Un valor numérico vacío** (`<v></v>`) se lee como una celda vacía en lugar de `#NUM!`. Un `<col>` sin ancho ya no se lee como `NaN`, y un borde con `style="none"` ya no se lee como un borde.
- **`rowHeights`, `hiddenRows` y `hiddenColumns`** aparecen en `ParsedSheet` solo en los archivos que los tienen.

### Workbook

- **Un ciclo de ida y vuelta ahora conserva** los bordes por lado, las alturas de fila, las filas ocultas y las columnas ocultas, los estilos de celda de fecha y los formatos de fecha personalizados, las siete celdas de error clásicas y el texto que empieza con `=`. Un archivo que dependía de que al guardar se mostraran las filas o columnas ocultas ahora las mantiene ocultas.
- **Los errores cargados y el texto con `=` conservan su tipo al guardar.** `getCellValue` y `getSheetData` siguen devolviendo las mismas cadenas (`'#N/A'`, `'=== Summary ==='`), así que el código que solo lee valores no se ve afectado. Volver a escribir esas cadenas con `setCellValue` o `addSheet` guarda `'=== Summary ==='` como fórmula y `'#N/A'` como texto, así que envuélvelas en `{ text }` y `{ error }`. Antes, un `=== Summary ===` cargado se guardaba como fórmula.
- **`removeAutoFilter` muestra las filas bajo el rango eliminado.** Las filas ocultas por un filtro siguen ocultas cuando cargas y guardas, porque los criterios no se conservan.
- **Los resultados de fórmula en caché no se conservan.** Las fórmulas siguen cargándose como cadenas `'=…'` y se guardan sin resultado.
- **Novedad:** `setRowHeight`, `getRowHeight`, `setRowHidden`, `isRowHidden`, `setColumnHidden` y `isColumnHidden`.

### Novedades y recomendaciones

Celdas `{ formula, result? }`, `{ text }` y `{ error }`; bordes por lado con 13 estilos de línea y colores RGB; estilos en celdas `Date`; `rowHeights`, `hiddenRows` y `hiddenColumns` en `ExcelWriter`, el escritor de streaming y `Workbook`. Prefiere `{ formula }` y `{ text }` en el código nuevo: la abreviatura de cadena `'='` sigue siendo una fórmula durante todo 1.x, y `{ text }` es la forma de guardar texto que empieza con `=`. Los tipos exportados `Font`, `Fill`, `Border`, `ExcelStyle` y `CellAlignment` ya no los usa la biblioteca y se eliminarán en 2.0.

## De 1.4 a 1.5

### Llamadas que ahora lanzan un error

Los escritores producían un archivo con estas entradas. Un archivo que Excel repara, o que otros lectores rechazan, es peor que un error en la llamada.

| Entrada | Qué hacer |
| --- | --- |
| Un nombre de hoja de más de 31 caracteres, con `\ / ? * [ ] :`, un carácter de control o un apóstrofo en cualquiera de los extremos, o dos nombres iguales sin distinguir mayúsculas de minúsculas | Acórtalo o cámbiale el nombre. |
| Un libro sin hojas, o con todas las hojas ocultas | Agrega una hoja, o deja una visible. |
| Un color que no es `#RGB`, `#RRGGBB` ni `#AARRGGBB` (`red`, `rgb(…)`) | Conviértelo a hexadecimal. |
| `NaN`, `Infinity` o un `Date` no válido en una celda | Reemplázalo antes de escribir, por ejemplo `Number.isFinite(value) ? value : null`. El error nombra la celda. |

Esto se aplica a `ExcelWriter`, `createExcelWorkbookStream`, `Workbook.toBuffer`/`toBlob` y a `Workbook.addSheet`, que ahora también compara los nombres sin distinguir mayúsculas de minúsculas.

Un libro cargado con `Workbook.fromBuffer` conserva los nombres que tenía. Si alguno tiene más de 31 caracteres, la carga funciona y guardar lanza un error; llama primero a `workbook.renameSheet(from, to)`. Las fórmulas que mencionan el nombre anterior no se reescriben.

El lector lanza un error para una celda o fila fuera de la cuadrícula de Excel (más allá de la columna `XFD` o de la fila 1,048,576) y para un libro que necesita más de 5,000,000 celdas vacías de relleno para mantener rectangulares las filas. Ambos casos corrían el riesgo de agotar la memoria.

### Valores que devuelve el lector

- **Las celdas de error** (`#DIV/0!`) se leían como el número `NaN`. Ahora se leen como `type: 'error'` con el texto como valor, así que un `switch` sobre `ParsedCell['type']` necesita un caso para `'error'`. Un número que no es finito (`1e999`) se lee como el error `#NUM!`.
- **Sin conversión numérica.** Los valores de atributos y de texto se devuelven tal como están escritos. Un formato de número `0.00` se leía como `0`, y un nombre de hoja o un título `007` como `7`. Si dependías de un número ahí, conviértelo tú mismo.
- **`ParsedSheet.validations`** contiene reglas completas (`type`, `operator`, `formula1`, `formula2`, `allowBlank`) en lugar de `{ range, options }`. `options` sigue conteniendo los valores de una lista en línea. Para los demás tipos ahora es la primera fórmula con sus comillas intactas (antes se quitaban), y las reglas de tipo `none` ya no se devuelven.
- **`ParsedSheet.state`** es `'hidden'` o `'veryHidden'` para las hojas ocultas y no existe en los demás casos.
- **`DataValidationType`** incorpora `'time'` y `'custom'`; un `switch` exhaustivo necesita ambos.

`ParsedSheet.data` no cambia: las filas presentes en el archivo, en orden. Usa `cell.rowIndex` para la posición.

### Workbook

- **Las filas se quedan donde estaban.** Cargar un archivo con filas en blanco y guardarlo movía hacia arriba todas las filas posteriores mientras los estilos y las celdas combinadas se quedaban en su sitio. Ahora las filas conservan su posición.
- **`getSheetData(name)` puede tener huecos.** Para un archivo con filas en blanco es disperso: `length` es la última fila más uno, `for...of` produce `undefined` para una fila que falta y `JSON.stringify` escribe `null`. Usa bucles por índice con `?.`, o `forEach`, que omite los huecos. `getCellValue(sheet, row, col)` ahora coincide con la fila del archivo.
- **Las validaciones, los formatos de número y las hojas ocultas sobreviven al guardar.** Los mensajes de validación y el estilo de error no; consulta la nota sobre el ciclo de ida y vuelta en el README.
- **Novedad:** `renameSheet`, `getSheetState`, `setSheetState`.

### Empaquetado

- El mapa `exports` da `index.d.mts` a quienes importan con ESM e `index.d.ts` a quienes importan con CommonJS, exporta `./package.json`, y el paquete declara `"sideEffects": false` y `"type": "commonjs"`.
- Las exportaciones de bajo nivel (`generate*Xml`, los ayudantes de zip) ya no se describen como estables.

### Salida y tamaño

- El escritor de streaming comprime con deflate cada entrada del zip, así que SheetJS 0.18.5 y el `ZipInputStream` de Java pueden leer su salida.
- `ExcelWriter` crece de 12.0 a 12.5 KB min+gzip y `ExcelReader` de 27.6 a 28.2 KB.
