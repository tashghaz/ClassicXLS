import Foundation

/// Errors the writer can throw.
public enum XLSWriteError: Error {
    case emptySheetName
    case invalidGrid(expectedWidth: Int, gotRowIndex: Int, gotWidth: Int)
    case notImplemented(String)   // <- temporary, until we add the encoder in Step 1
}

/// A simple table you want to write into a single worksheet.
/// - `name`: worksheet/tab name
/// - `headers`: row 0, starting at column B (we skip column A to match your reader)
/// - `rows`: data rows (each MUST match headers.count)
public struct XLSWriteSheet {
    public let name: String
    public let headers: [String]
    public let rows: [[String]]

    public init(name: String, headers: [String], rows: [[String]]) {
        self.name = name
        self.headers = headers
        self.rows = rows
    }
}

/// Public entry point for writing a legacy .xls file.
public enum XLSWriter {

    /// Write `sheet` into a new .xls file at `url`.
    /// Step 0: this only validates inputs and throws `.notImplemented`.
    /// Step 1 will add the real encoder behind this API.
    public static func write(sheet: XLSWriteSheet, to url: URL) throws {
        // 0) Validate name
        guard !sheet.name.isEmpty else {
            throw XLSWriteError.emptySheetName
        }

        // 1) Validate widths (every row must match headers.count)
        let expectedWidth = sheet.headers.count
        for (rowIndex, row) in sheet.rows.enumerated() where row.count != expectedWidth {
            throw XLSWriteError.invalidGrid(expectedWidth: expectedWidth, gotRowIndex: rowIndex, gotWidth: row.count)
        }

        // 2) Placeholder until Step 1 (actual BIFF/OLE encoding)
        throw XLSWriteError.notImplemented("Writer encoder will be added in Step 1.")
    }
}