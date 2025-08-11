# ClassicXLS

[![Swift](https://img.shields.io/badge/Swift-5.9-orange.svg)](https://swift.org)
[![SPM Compatible](https://img.shields.io/badge/SPM-compatible-green.svg)](https://swift.org/package-manager/)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

**ClassicXLS** is a pure Swift library for reading legacy Excel `.xls` (BIFF) files on macOS and iOS.  
No Excel, COM, or external runtime required. Works fully offline, MIT licensed.

---

## ✨ Features

- 📂 Read `.xls` files from Excel 97–2003 (BIFF8 format)
- 🛡 Safe parsing — no unsafe pointer crashes
- 🖥 Works on **macOS** and **iOS**
- 📦 Installable via **Swift Package Manager**
- 📜 MIT licensed — free for personal and commercial use

---

## 📦 Installation

### Swift Package Manager (SPM)

In Xcode:

1. Go to **File** → **Add Packages…**
2. Enter the URL: 
https://github.com/tashghaz/ClassicXLS

3. Choose a release tag, and add to your target.

---

## 🚀 Usage

```swift
import ClassicXLS

do {
 let fileURL = URL(fileURLWithPath: "/path/to/workbook.xls")
 let workbook = try XLSReader.read(url: fileURL)

 print("Workbook contains \(workbook.sheets.count) sheets:")
 for sheet in workbook.sheets {
     print("- \(sheet.name) (\(sheet.grid.count) rows)")

     // Access cell values
     if let row = sheet.grid[0] {
         for (_, cell) in row {
             print("  Row \(cell.row), Col \(cell.col): \(cell.value)")
         }
     }
 }
} catch {
 print("❌ Failed to read XLS: \(error)")
}

```

---
## 📄 License

This project is licensed under the MIT License – see the [LICENSE](LICENSE) file for details.
---
