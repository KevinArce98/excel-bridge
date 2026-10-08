---
title: Límites y alternativas
description: Lo que excel-bridge no hace, los casos que lee mal, dónde se ha probado y cuándo conviene otra biblioteca.
group: Referencia
groupOrder: 3
order: 3
---

# Límites y alternativas

excel-bridge escribe informes y los lee de vuelta. No lo cubre todo, y esta página dice dónde se detiene.

## Cuándo conviene otra biblioteca

- **Imágenes, gráficos, comentarios, tablas de Excel, tablas dinámicas, configuración de impresión, texto enriquecido, bordes diagonales.** Ninguno de estos está soportado. ExcelJS y hucre cubren muchos de ellos.
- **Editar un archivo y conservar todo lo que no tocaste**, como las plantillas con gráficos o macros. `Workbook` reconstruye el archivo a partir de lo que modela, así que todo lo demás se descarta. `openXlsx` y `saveXlsx` de hucre conservan las partes que no modelan.
- **Leer archivos muy grandes sin cargarlos completos.** `ExcelReader` no tiene modo de streaming. hucre tiene `streamXlsxRows`.
- **CSV, ODS o los antiguos `.xls` y `.xlsb`.** Usa hucre o SheetJS.
- **Importaciones validadas con un esquema** que asignan filas a objetos con tipo y con listas de errores. Usa read-excel-file.

## Dónde se ha probado

| Entorno | Soporte |
| --- | --- |
| Node.js | `^20.19.0`, `^22.13.0` o `>=24`, igual que `engines`. La suite de pruebas se ejecuta en Node 20, 22 y 24. |
| Navegadores | No se ejecutan en CI. El código apunta a ES2022 y necesita `File` y `Blob`. `downloadXlsx` se prueba con un `document` simulado. |
| Formatos de módulo | ESM (`import`) y CommonJS (`require`). |
| Bun y Deno | El paquete empaquetado se somete a una prueba de humo en Node 22, Bun 1.x y Deno 2.x en CI. |
| Excel | Escrito para Excel 2016 y posteriores, pero CI no abre el resultado en Excel, LibreOffice ni Google Sheets. El lector se prueba con archivos escritos por ExcelJS, SheetJS y hucre. |

## Escritura

- **Las fórmulas se recalculan al abrir.** El libro le pide a Excel que recalcule todas las fórmulas al cargarlo. Una fórmula sin `result` no lleva ningún valor almacenado, así que un lector que no calcula la muestra vacía: SheetJS con las opciones predeterminadas omite la celda, y openpyxl con `data_only` devuelve `None`.
- **Las fechas son valores de hora local de reloj.** `new Date(2024, 0, 15)` se escribe como el 15 de enero en cualquier zona horaria, con el formato de fecha corta integrado, a menos que el estilo de la celda tenga un `numberFormat` de fecha. Una fecha a medianoche UTC cae en el día anterior con desfases negativos, y una hora local dentro de un salto por horario de verano se adelanta.
- **Los bordes cubren los cuatro lados**, con 13 estilos de línea y colores RGB. No hay bordes diagonales ni bordes en los formatos condicionales.
- **Valores predeterminados de diseño.** Una fila sin altura usa el valor predeterminado de Excel. Los valores predeterminados de la hoja (altura de fila y ancho de columna predeterminados, niveles de esquema) no se escriben ni se conservan.
- **Solo rangos de AutoFilter.** Los criterios de filtro y el estado de ordenación no se escriben ni se leen.
- **Esquemas de hipervínculo.** El escritor acepta `http:`, `https:`, `mailto:` y ubicaciones dentro del libro.
- **El escritor de streaming** no admite `autoWidth`, `validations`, `conditionalFormats`, `sharedStrings` ni un `state` de hoja, y siempre escribe las cadenas en línea.

## Lectura

El lector no está reforzado para archivos no confiables más allá de sus [límites](../guide/errors/). Mantiene todo el archivo en memoria.

Un caso todavía no se lee correctamente: un seguidor de fórmula compartida se lee como su valor en caché, sin fórmula.

## Lo que no afirma

- No es el más pequeño en cada fila de la comparación. Consulta [Rendimiento y tamaño](../performance/).
- No es el más rápido: hucre escribe la carga de trabajo del benchmark más rápido (unos 390 ms frente a 560 ms).
- Nada de lo que se dice aquí afirma que un archivo se haya probado en Excel, LibreOffice o Google Sheets, porque CI no abre ninguno.
