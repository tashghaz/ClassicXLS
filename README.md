# ClassicXLS

A pure Swift library to read legacy **.xls** (BIFF5/BIFF8) files — no third-party dependencies.

## Status
WIP: parsing OLE/CFB container and BIFF records step by step. Early API:

```swift
import ClassicXLS
let workbook = try XLSReader.read(url: fileURL)
for sheet in workbook.sheets {
    print("Sheet:", sheet.name)
}
