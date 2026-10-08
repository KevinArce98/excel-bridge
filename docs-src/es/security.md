---
title: Seguridad
description: Cómo reportar una vulnerabilidad y qué hacer antes de leer un .xlsx que no creaste tú.
group: Referencia
groupOrder: 3
order: 5
---

# Seguridad

## Reportar una vulnerabilidad

Reporta las vulnerabilidades en privado mediante los [avisos de seguridad de GitHub](https://github.com/KevinArce98/excel-bridge/security/advisories/new). La [política de seguridad](https://github.com/KevinArce98/excel-bridge/blob/main/SECURITY.md) enumera qué está dentro del alcance. Por favor, no abras un issue público para reportar una vulnerabilidad.

## Antes de escribir datos de usuarios

- **Las cadenas no pueden convertirse en fórmulas.** Una cadena simple siempre es texto, así que `=HYPERLINK(...)` en el nombre de un cliente se escribe como texto. Solo `{ formula }` escribe una fórmula en una celda, así que nunca le pases entrada no confiable. La `expression` de un formato condicional y la `formula1` de una validación `custom` también son fórmulas.
- **Los vínculos se comprueban.** El escritor solo acepta URL `http:`, `https:` y `mailto:` y ubicaciones dentro del libro, y lanza un error para cualquier otra cosa.

## Antes de leer un archivo que no creaste tú

- **Los límites son un mínimo, no una garantía.** `maxPartBytes`, `maxTotalBytes` y `maxSheets` rechazan los archivos grandes antes de descomprimir las hojas. `maxCells` se comprueba mientras se analiza una hoja, después de descomprimir su XML. Un archivo por debajo de todos los valores predeterminados aún puede usar mucha memoria: leer una hoja de 205 MiB llegó a entre 2.1 y 2.5 GB de memoria residente, unas 10 a 12 veces el XML de la hoja.
- **Baja `maxSheets`** para cargas públicas. Por defecto es 1,000. Un archivo puede listar muchas hojas que apunten a una misma parte grande, que se descomprime una vez pero se analiza una vez por hoja.
- **Limita el tamaño de la carga** antes de analizar, y define los límites según lo que necesite tu producto.
- **Analiza las cargas no confiables en un worker o en un proceso separado** que se pueda reiniciar.
- **Comprueba los esquemas de hipervínculo** antes de mostrar un vínculo leído de un archivo. `ExcelReader` devuelve los vínculos tal como están guardados.
- **El lector de XML rechaza `DOCTYPE`**, así que un archivo no puede expandir entidades.

Consulta [Maneja errores y límites](../guide/errors/) para ver las opciones.
