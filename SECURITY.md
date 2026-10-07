# Security policy

## Supported versions

Security fixes go into the latest release of `excel-bridge`.

## Reporting a vulnerability

Use [GitHub private vulnerability reporting](https://github.com/KevinArce98/excel-bridge/security/advisories/new)
to send the report. Please do not open a public issue for a vulnerability.

Include the version, a minimal file or input that reproduces the problem, and what you observed.

## Scope and known limits

`excel-bridge` parses `.xlsx` files in the process that calls it. The reader:

- rejects cell and row references outside Excel's grid, clamps `<col>` ranges, inflates only the
  parts it uses and stops when a workbook needs more than 5,000,000 empty padding cells;
- does not limit the size of a decompressed worksheet, and holds the whole workbook in memory, so a
  very large legitimate or crafted sheet can exhaust memory.

Cap the size of uploads before parsing them, and parse untrusted files in a worker or a separate
process. The reader returns hyperlink targets as stored (a `javascript:` target comes back
unchanged; `Workbook` drops targets the writer does not accept) and returns text unsanitized, so
escape it before putting it in HTML.

The writers do not escape text that starts with `=`: such a string becomes a live formula in the
file, so do not pass user-controlled text that starts with `=` to them. The same happens on a
`Workbook` load and save, because a text cell that starts with `=` in a loaded file is saved as a
formula.
