---
title: Security
description: How to report a vulnerability, and what to do before you read an .xlsx you did not make.
group: Reference
groupOrder: 3
order: 5
---

# Security

## Report a vulnerability

Report vulnerabilities privately through [GitHub security advisories](https://github.com/KevinArce98/excel-bridge/security/advisories/new). The [security policy](https://github.com/KevinArce98/excel-bridge/blob/main/SECURITY.md) lists what is in scope. Please do not open a public issue for a vulnerability.

## Before you write user data

- **Strings cannot become formulas.** A plain string is always text, so `=HYPERLINK(...)` in a customer name is written as text. Only `{ formula }` writes a formula in a cell, so never pass it untrusted input. The `expression` of a conditional format and the `formula1` of a `custom` validation are formulas too.
- **Links are checked.** The writer accepts only `http:`, `https:` and `mailto:` URLs and locations inside the workbook, and throws for anything else.

## Before you read a file you did not make

- **The limits are a floor, not a guarantee.** `maxPartBytes`, `maxTotalBytes` and `maxSheets` refuse large files before the sheets are inflated. `maxCells` is checked while a sheet is parsed, after its XML is inflated. A file under every default can still use a lot of memory: reading a 205 MiB sheet peaked at 2.1 to 2.5 GB of resident memory, about 10 to 12 times the sheet XML.
- **Lower `maxSheets`** for public uploads. It defaults to 1,000. A file can list many sheets that point at one large part, which is inflated once but parsed once per sheet.
- **Cap the upload size** before you parse, and set the limits to what your product needs.
- **Parse untrusted uploads in a worker or a separate process** that can be restarted.
- **Check hyperlink schemes** before you render a link read from a file. `ExcelReader` returns links as stored.
- **The XML reader rejects `DOCTYPE`**, so a file cannot expand entities.

See [Handle errors and limits](../guide/errors/) for the options.
