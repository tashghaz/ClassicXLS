import Foundation

public enum XLSWriteError: Error {
    case emptySheetName
    case invalidGrid(expectedWidth: Int, gotRowIndex: Int, gotWidth: Int)
}

public struct XLSWriteSheet {
    public let name: String
    public let headers: [String]
    public let rows: [[String]]
    public init(name: String, headers: [String], rows: [[String]]) {
        self.name = name; self.headers = headers; self.rows = rows
    }
}

public enum XLSWriter {
    public static func write(sheet: XLSWriteSheet, to url: URL) throws {
        guard !sheet.name.isEmpty else { throw XLSWriteError.emptySheetName }
        let w = sheet.headers.count
        for (i, r) in sheet.rows.enumerated() where r.count != w {
            throw XLSWriteError.invalidGrid(expectedWidth: w, gotRowIndex: i, gotWidth: r.count)
        }

        let worksheet = BIFFWorksheetBuilder.makeWorksheetStream(
            sheetName: sheet.name, headers: sheet.headers, rows: sheet.rows
        )
        let workbook = BIFFWorkbookBuilder.buildWorkbook(
            sheetName: sheet.name, worksheetData: worksheet
        )
        try CFBWriter.writeSingleStream(streamName: "Book", stream: workbook, to: url)
    }
}