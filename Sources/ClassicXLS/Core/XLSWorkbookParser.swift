
import Foundation

// Workbook + sheet parsing (to be implemented in Steps 3–4)
struct XLSWorkbookParser {
    static func parse(workbookStream: Data) throws -> XLSWorkbook {
        // Will parse SST, BOUNDSHEET, then read sheets
        return XLSWorkbook(sheets: [])
    }
}
