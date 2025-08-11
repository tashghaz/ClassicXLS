
import Foundation

public enum XLSReadError: Error, CustomStringConvertible {
    case notXLS
    case workbookStreamMissing
    case parseError(String)

    public var description: String {
        switch self {
        case .notXLS: return "File is not an OLE2/CFB .xls workbook"
        case .workbookStreamMissing: return "Workbook/Book stream not found"
        case .parseError(let m): return "Parse error: \(m)"
        }
    }
}

public enum XLSReader {
    /// Reads a legacy .xls file and returns a parsed workbook.
    /// For now this is a placeholder; real parsing arrives in Step 2–4.
    public static func read(url: URL) throws -> XLSWorkbook {
        // Step 2 will: open OLE/CFB, extract "Workbook" or "Book" stream
        // Step 3 will: parse BIFF globals (SST, BOUNDSHEET)
        // Step 4 will: parse sheet cells (NUMBER, RK, LABELSST, etc.)
        // Temporary stub so the package builds:
        return XLSWorkbook(sheets: [])
    }
}
