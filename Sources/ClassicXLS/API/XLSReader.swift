
import Foundation

// MARK: - Errors

public enum XLSReadError: Error, CustomStringConvertible {
    case notXLS
    case workbookStreamMissing
    case parseError(String)

    public var description: String {
        switch self {
        case .notXLS:
            return "File is not an OLE2/CFB .xls workbook"
        case .workbookStreamMissing:
            return "Workbook/Book stream not found"
        case .parseError(let m):
            return "Parse error: \(m)"
        }
    }
}

// Handy for tests
extension XLSReadError: Equatable {
    public static func == (lhs: XLSReadError, rhs: XLSReadError) -> Bool {
        switch (lhs, rhs) {
        case (.notXLS, .notXLS), (.workbookStreamMissing, .workbookStreamMissing): return true
        case (.parseError, .parseError): return true
        default: return false
        }
    }
}

// MARK: - Public Facade

public enum XLSReader {
    /// Opens a legacy .xls file.
    /// Step 2: extracts the OLE "Workbook"/"Book" BIFF stream (parsing comes next).
    public static func read(url: URL) throws -> XLSWorkbook {
        let cfb = try CFBFile(fileURL: url)

        // Try "Workbook" first, then older "Book"
        let wbStream: Data
        if let s = try? cfb.stream(named: "Workbook") {
            wbStream = s
        } else if let s = try? cfb.stream(named: "Book") {
            wbStream = s
        } else {
            throw XLSReadError.workbookStreamMissing
        }

        // TODO (Step 3/4): parse `wbStream` (SST, BOUNDSHEET, then sheets)
        _ = wbStream
        #if DEBUG
        print("ClassicXLS: extracted Workbook stream, bytes=\(wbStream.count)")
        #endif

        // Temporary stub so the package builds/links with callers
        return XLSWorkbook(sheets: [])
    }
}
