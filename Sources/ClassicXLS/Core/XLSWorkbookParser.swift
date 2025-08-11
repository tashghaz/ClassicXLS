//
//  File.swift
//  ClassicXLS
//
//  Created by Artashes Ghazaryan on 8/11/25.
//

import Foundation

// Workbook + sheet parsing (to be implemented in Steps 3–4)
struct XLSWorkbookParser {
    static func parse(workbookStream: Data) throws -> XLSWorkbook {
        // Will parse SST, BOUNDSHEET, then read sheets
        return XLSWorkbook(sheets: [])
    }
}
