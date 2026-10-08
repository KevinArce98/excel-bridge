---
title: Rendimiento y tamaño
description: El tamaño del paquete de cada punto de entrada junto a otras bibliotecas, la velocidad de escritura y de lectura, y cómo se midió cada número.
group: Referencia
groupOrder: 3
order: 2
---

# Rendimiento y tamaño

Cada número de esta página sale de un script del repositorio, y el método está junto a él. Ejecuta los scripts para comprobarlos en tu propia máquina.

## Tamaño del paquete

Los empaquetadores conservan las exportaciones que importas y descartan el resto. Esto es lo que suma cada punto de entrada a un paquete para navegador:

| Importar desde `excel-bridge` | min+gzip |
| --- | ---: |
| `createExcelWorkbookStream` | 12.3 KB |
| `ExcelWriter` | 12.8 KB |
| `ExcelReader` | 9.5 KB |
| `Workbook` (lector + escritor) | 21.5 KB |
| `ExcelBridge` (objeto de conveniencia) | 21.6 KB |
| Todo | 25.1 KB |

La misma medición para otras bibliotecas (2026-10-07):

| Biblioteca | Importación | min+gzip |
| --- | --- | ---: |
| hucre 1.2.0 | `writeXlsx` | 41.6 KB |
| hucre 1.2.0 | `readXlsx` | 41.1 KB |
| hucre 1.2.0 | `XlsxStreamWriter` (desde `hucre/xlsx`) | 12.1 KB |
| hucre 1.2.0 | `streamXlsxRows` (desde `hucre/xlsx`) | 17.4 KB |
| @mitresthen/excelents 1.0.1 | toda la entrada principal, lector y escritor | 12.2 KB |
| read-excel-file 9.3.10 | `read-excel-file/browser`, solo lectura | 16.7 KB |
| write-excel-file 4.1.1 | `write-excel-file/universal`, solo escritura | 19.8 KB |
| SheetJS (`xlsx` 0.18.5 en npm) | `utils` + `write` | 95.8 KB |
| ExcelJS 4.4.0 | compilación de navegador predeterminada, no admite tree-shaking | 272.1 KB |

**excel-bridge no es el más pequeño en cada fila.** `ExcelWriter` escribe estilos, bordes, formato condicional, validación de datos, autofiltros e hipervínculos en 12.8 KB. Algunas bibliotecas son más pequeñas para una sola tarea, o renuncian a una función para lograrlo: la entrada de `excelents` no tiene formato condicional en 1.0.1, y lee y escribe en menos espacio que todo excel-bridge. Compara lo que obtienes por los bytes, no solo los bytes.

El tree-shaking usa la compilación ESM, que los empaquetadores eligen para `import`. `require('excel-bridge')` carga la compilación CommonJS completa. `ExcelBridge` pesa 21.6 KB frente a 12.8 KB de `ExcelWriter`, porque los empaquetadores conservan un objeto completo: incluso una sola llamada a `ExcelBridge.write` incluye también el lector.

**Cómo se midieron estos números**

- **Herramientas:** esbuild 0.28.2 (`--bundle --minify --platform=browser --format=esm`). Las filas de excel-bridge se midieron el 2026-10-08 y las de las demás bibliotecas el 2026-10-07.
- **Entrada:** cada fila empaqueta un archivo de una línea `export { … } from '<package>'` y luego comprime la salida con gzip mediante zlib de Node en el nivel predeterminado. 1 KB = 1,000 bytes.
- **Fuentes:** las filas de excel-bridge salen de la compilación de este repositorio. Las demás bibliotecas se instalaron en un directorio temporal.
- **Margen:** las implementaciones de gzip difieren en cerca de 1%. El `gzip` de macOS da un resultado un poco más pequeño.
- **No medido:** las compilaciones más recientes de SheetJS Community Edition, que se distribuyen desde el CDN de SheetJS.
- **Reproducir:** ejecuta `pnpm run size`. CI ejecuta `pnpm run size:check`, que falla cuando una fila de excel-bridge se aleja más de 100 bytes del tamaño medido.

## Cómo se compara

| | **excel-bridge** | ExcelJS | SheetJS (`xlsx` de la comunidad) |
| --- | :---: | :---: | :---: |
| Leer `.xlsx` | ✅ | ✅ | ✅ |
| Escribir `.xlsx` | ✅ | ✅ | ✅ |
| Estilos de celda (color, fuente, bordes por lado) | ✅ | ✅ | ⚠️ Edición Pro |
| Formato condicional | ✅ | ✅ | ⚠️ Edición Pro |
| Fórmulas | ✅ | ✅ | ✅ |
| Celdas combinadas | ✅ | ✅ | ✅ |
| Paneles inmovilizados | ✅ | ✅ | ❌ no está en `xlsx` 0.18.5 |
| Escritor de streaming | ✅ | ✅ | ⚠️ Edición Pro |
| Tipos de TypeScript de primera clase | ✅ | ✅ | ✅ |
| ESM **y** CJS, con tree-shaking | ✅ | ⚠️ orientado a CJS | ✅ |
| Dependencias directas en tiempo de ejecución ² | 1 | 9 | 7 |
| Tamaño del paquete para escribir un archivo ¹ | **12.8 KB** | 272.1 KB | 95.8 KB |

¹ Código minificado y comprimido con gzip que un paquete de navegador necesita para escribir un `.xlsx`: `ExcelWriter`, la compilación de navegador predeterminada de ExcelJS (no admite tree-shaking) y `utils` + `write` de SheetJS, de `xlsx@0.18.5` en npm, empaquetados con esbuild. Bundlephobia mide cada paquete completo con su propia cadena de herramientas, así que sus números difieren: [excel-bridge](https://bundlephobia.com/package/excel-bridge), [exceljs](https://bundlephobia.com/package/exceljs) y [xlsx](https://bundlephobia.com/package/xlsx).

² Declaradas en el `package.json` de cada una (`exceljs@4.4.0`, `xlsx@0.18.5` en npm), comprobadas el 2026-10-07. `xlsx@0.18.5` es la última versión publicada en npm y tiene avisos de seguridad conocidos.

## Velocidad de escritura

Escribir 50,000 filas × 10 columnas (Node 24.19, Apple M4, 2026-10-08). Cada tiempo es la mediana de 5 ejecuciones dentro de una invocación, tomada sobre tres invocaciones:

| Biblioteca | Tiempo de escritura | Tamaño de salida |
| --- | ---: | ---: |
| **excel-bridge** | **560 ms** | **2.41 MiB** |
| excel-bridge (streaming) | 580 ms | 2.48 MiB |
| hucre | 390 ms | 2.79 MiB |
| xlsx / SheetJS (predeterminado, sin compresión) | 510 ms | 18.23 MiB |
| xlsx / SheetJS (`compression: true`) | 580 ms | 6.26 MiB |
| exceljs | 1430 ms | 2.82 MiB |

- **Frente a ExcelJS:** alrededor de 2.6× más rápido.
- **Frente a hucre:** hucre escribe esta carga de trabajo más rápido, en unos 390 ms.
- **Frente a SheetJS:** un tiempo similar, 510 ms con su salida predeterminada sin comprimir y 580 ms con `compression: true`. Su archivo es unas 7.5× más grande de forma predeterminada y unas 2.6× más grande con compresión.

Los tiempos varían 10% o más entre ejecuciones y según la máquina, y las proporciones se mantienen mejor que los tiempos absolutos. Reprodúcelos con `pnpm run bench` tras instalar las otras bibliotecas (consulta el README de benchmarks).

## Velocidad de lectura

Lectura de un archivo que escribió excel-bridge, de 50,000 filas × 10 columnas y de 200,000 filas × 10 columnas (Node 24.19, Apple M4, 2026-10-08, mediana de 7 y 3 ejecuciones):

| Archivo | 1.6.0 | 2.0 | Pico de memoria 1.6.0 | Pico de memoria 2.0 |
| --- | ---: | ---: | ---: | ---: |
| 50,000 × 10 | 1503 ms | 270 ms | 859 MiB | 451 MiB |
| 200,000 × 10 | 6521 ms | 1435 ms | 2679 MiB | 1611 MiB |

La versión 2.0 lee entre 4.5 y 5.6 veces más rápido y alcanza un pico de entre la mitad y dos tercios de la memoria. El lector analiza el XML con su propio tokenizador pequeño en lugar de un analizador de uso general, y esa es también la razón por la que `ExcelReader` pasó de 28.7 a 9.5 KB.

El pico de memoria es el del proceso del benchmark, que también construye el archivo que se lee. El lector mantiene todo el archivo en memoria y no tiene modo de streaming. Reproduce las cifras con `node --expose-gc benchmarks/bench-read.mjs`.
