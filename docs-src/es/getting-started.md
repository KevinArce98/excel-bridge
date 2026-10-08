---
title: Primeros pasos
description: Instala excel-bridge, escribe un .xlsx con estilos y léelo de vuelta, en el navegador y en Node.js.
group: Inicio
groupOrder: 1
order: 1
---

# Primeros pasos

Escribe una hoja de cálculo con estilos y léela de vuelta, en el navegador o en Node.js, con un solo paquete pequeño. Importa una clase y tu empaquetador incluye solo esa parte: el escritor suma 12.9 KB min+gzip.

## Instalación

```bash tab="npm"
npm install excel-bridge
```

```bash tab="pnpm"
pnpm add excel-bridge
```

```bash tab="yarn"
yarn add excel-bridge
```

```bash tab="bun"
bun add excel-bridge
```

excel-bridge necesita Node.js 20.19, 22.13 o 24 y posteriores, o un navegador moderno. Está escrito en TypeScript e incluye compilaciones ESM y CommonJS con tipos.

## Escribir un libro

Cada hoja es un objeto con `data` y, de forma opcional, `styles` y `options`. Los estilos usan como clave `"<row>-<col>"`, con índices que empiezan en cero.

```ts title="write.ts"
import fs from 'node:fs';
import { ExcelWriter } from 'excel-bridge';

const header = { bold: true, background: '#4472C4', color: '#FFFFFF' };

const sheet = {
  data: [
    ['Name', 'Age', 'City'],
    ['John', 25, 'New York'],
    ['Jane', 30, 'Los Angeles'],
  ],
  styles: { '0-0': header, '0-1': header, '0-2': header },
  options: { name: 'People', freezePane: { row: 1 }, autoWidth: true },
};

const writer = new ExcelWriter();

fs.writeFileSync('people.xlsx', writer.createWorkbookBuffer([sheet]));
```

En el navegador, `createWorkbook([sheet])` devuelve un `Blob`. Pásalo a `downloadXlsx(blob, 'people')` para iniciar la descarga.

> [!NOTE]
> `createWorkbookBuffer` devuelve un `Uint8Array` simple, no un `Buffer`. Express 4 necesita `res.send(Buffer.from(bytes))`.

## Leerlo de vuelta

```ts title="read.ts"
import fs from 'node:fs';
import { ExcelReader } from 'excel-bridge';

const workbook = new ExcelReader().parseFromBuffer(fs.readFileSync('people.xlsx'));

workbook.sheets[0].data.forEach(row => {
  row.forEach(cell => console.log(cell.coordinate, cell.type, cell.value));
});
```

Cada celda lleva su `type`. Las fechas vuelven como `Date`, y una celda con fórmula guarda su fórmula en `formula`. En el navegador, `await reader.parseFromFile(file)` recibe el `File` de un `<input type="file">`.

## Editar un archivo

`Workbook` carga un archivo, te permite modificarlo y lo guarda de nuevo.

```ts title="edit.ts"
import fs from 'node:fs';
import { Workbook } from 'excel-bridge';

const workbook = Workbook.fromBuffer(fs.readFileSync('people.xlsx'));
workbook.setCellValue('People', 1, 1, 26);
fs.writeFileSync('people.xlsx', workbook.toBuffer());
```

> [!LIMIT]
> `Workbook` reconstruye el archivo a partir de lo que modela, así que las imágenes, los gráficos y los comentarios de un archivo cargado se descartan. Consulta [Edita un libro](../guide/workbook/).

## Siguientes pasos

- [Escribe valores y fórmulas](../guide/values/): qué puede contener una celda y cómo escribir una fórmula.
- [Da estilo a las celdas](../guide/styling/) y [Define el diseño de las hojas](../guide/layout/): bordes, anchos, celdas combinadas, filtros.
- [Lee un libro](../guide/reading/) y [Filas como objetos](../guide/objects/): celdas con tipo y filas con tipo.
- [Exporta archivos grandes con streaming](../guide/streaming/): archivos de un millón de filas y entrega de archivos.
- [Referencia de la API](../api/) y [Actualización](../upgrading/).
