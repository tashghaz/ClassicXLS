import Foundation

/// Errors the writer can throw.
public enum XLSWriteError: Error {
    case emptySheetName
    case invalidGrid(expectedWidth: Int, gotRowIndex: Int, gotWidth: Int)
    case notImplemented(String)
}

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
    public static func write(sheet: XLSWriteSheet, to url: URL) throws {
    // 0) Validate
    guard !sheet.name.isEmpty else {
        throw XLSWriteError.emptySheetName
    }
    let expectedWidth = sheet.headers.count
    for (rowIndex, row) in sheet.rows.enumerated() where row.count != expectedWidth {
        throw XLSWriteError.invalidGrid(expectedWidth: expectedWidth, gotRowIndex: rowIndex, gotWidth: row.count)
    }

    // 1) Build the worksheet stream (Step 1)
    let worksheetData = BIFFWorksheetBuilder.makeWorksheetStream(
        sheetName: sheet.name,
        headers: sheet.headers,
        rows: sheet.rows
    )

    // 2) Build the workbook stream (Step 2)
    let workbookData = BIFFWorkbookBuilder.makeWorkbookStream(
        sheetName: sheet.name,
        worksheetStream: worksheetData
    )

    // 3) Stop here in Step 2 (no OLE/CFB yet)
    //    We'll wrap `workbookData` in an OLE container in Step 3.
    throw XLSWriteError.notImplemented("Step 2 complete. Next: OLE container in Step 3.")
}