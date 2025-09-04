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
public static func write(sheet: XLSWriteSheet, to url: URL) throws {
    // 0) Validate
    guard !sheet.name.isEmpty else {
        throw XLSWriteError.emptySheetName
    }
    let expectedWidth = sheet.headers.count
    for (rowIndex, row) in sheet.rows.enumerated() where row.count != expectedWidth {
        throw XLSWriteError.invalidGrid(expectedWidth: expectedWidth, gotRowIndex: rowIndex, gotWidth: row.count)
    }

    // 1) Worksheet stream (Step 1)
    let worksheetData = BIFFWorksheetBuilder.makeWorksheetStream(
        sheetName: sheet.name,
        headers: sheet.headers,
        rows: sheet.rows
    )

    // 2) Workbook stream (Step 2)
    let workbookData = BIFFWorkbookBuilder.makeWorkbookStream(
        sheetName: sheet.name,
        worksheetStream: worksheetData
    )

    // 3) OLE/CFB container (Step 3) — write to disk as a real .xls
    try CFBWriter.writeSingleStream(streamName: "Book", stream: workbookData, to: url)
}
}