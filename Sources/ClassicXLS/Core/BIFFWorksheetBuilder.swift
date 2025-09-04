import Foundation

/// Builds a single BIFF5 **worksheet** stream as `Data`.
/// This does **not** build the workbook globals or OLE container.
/// Rows/columns are zero-based at the BIFF level:
/// - We place headers at row 0, **starting at column 1 (B)** to match your reader convention.
/// - We place data rows at **row 1..**, same columns.
enum BIFFWorksheetBuilder {

    // MARK: - Public API

    /// Convert headers + rows to a BIFF5 worksheet stream.
    static func makeWorksheetStream(sheetName: String,
                                    headers: [String],
                                    rows: [[String]]) -> Data {
        var worksheetData = Data()
        appendBOFWorksheet(to: &worksheetData)
        appendDimensions(headers: headers, rows: rows, to: &worksheetData)
        appendHeaderRow(headers, to: &worksheetData)
        appendDataRows(rows, to: &worksheetData)
        appendEOF(to: &worksheetData)
        return worksheetData
    }

    // MARK: - Record builders

    /// BOF (worksheet) — sid 0x0809, type 0x0010, BIFF5
    private static func appendBOFWorksheet(to data: inout Data) {
        var payload = Data()
        payload.append(le16(0x0500))  // BIFF5
        payload.append(le16(0x0010))  // type: worksheet
        payload.append(le16(0))       // build id
        payload.append(le16(0))       // build year
        payload.append(le32(0))       // file history flags
        payload.append(le32(0))       // lowest Excel ver
        data.append(record(sid: 0x0809, payload: payload))
    }

    /// DIMENSIONS — sid 0x0200
    /// rowMin,rowMax (exclusive), colMin,colMax (exclusive), flags
    private static func appendDimensions(headers: [String],
                                         rows: [[String]],
                                         to data: inout Data) {
        let firstRowIndex: UInt32 = 0               // header row
        let firstDataRowIndex: UInt32 = 1           // data starts after header
        let lastRowExclusive: UInt32 = firstDataRowIndex + UInt32(rows.count)

        let firstColumnIndex: UInt16 = 1            // start at column B
        let tableWidth: Int = max(headers.count, rows.first?.count ?? 0)
        let lastColumnExclusive: UInt16 = UInt16(firstColumnIndex + UInt16(tableWidth))

        var payload = Data()
        payload.append(le32(firstRowIndex))
        payload.append(le32(lastRowExclusive))
        payload.append(le16(firstColumnIndex))
        payload.append(le16(lastColumnExclusive))
        payload.append(le16(0)) // flags/reserved
        data.append(record(sid: 0x0200, payload: payload))
    }

    /// Header cells at row 0, columns B.. (1..)
    private static func appendHeaderRow(_ headers: [String], to data: inout Data) {
        let headerRowIndex = 0
        let firstColumnIndex = 1
        for (offset, text) in headers.enumerated() {
            let col = firstColumnIndex + offset
            data.append(labelCell(row: headerRowIndex, column: col, text: text))
        }
    }

    /// Data rows at row 1.., columns B.. (1..)
    private static func appendDataRows(_ rows: [[String]], to data: inout Data) {
        let firstDataRowIndex = 1
        let firstColumnIndex = 1
        for (rowOffset, rowValues) in rows.enumerated() {
            let rowIndex = firstDataRowIndex + rowOffset
            for (colOffset, cell) in rowValues.enumerated() {
                let colIndex = firstColumnIndex + colOffset
                if let numeric = Double(cell.replacingOccurrences(of: ",", with: ".")) {
                    data.append(numberCell(row: rowIndex, column: colIndex, value: numeric))
                } else {
                    data.append(labelCell(row: rowIndex, column: colIndex, text: cell))
                }
            }
        }
    }

    /// EOF — sid 0x000A
    private static func appendEOF(to data: inout Data) {
        data.append(record(sid: 0x000A, payload: Data()))
    }

    // MARK: - Cell record helpers

    /// NUMBER (sid 0x0203): row(2), col(2), xf(2), IEEE754 double (8)
    private static func numberCell(row: Int, column: Int, value: Double) -> Data {
        var payload = Data()
        payload.append(le16(UInt16(row)))
        payload.append(le16(UInt16(column)))
        payload.append(le16(0))                        // XF index
        payload.append(le64(value.bitPattern))         // raw 8 bytes of double
        return record(sid: 0x0203, payload: payload)
    }

    /// LABEL (sid 0x0204): row(2), col(2), xf(2), len(1), bytes (8-bit)
    /// BIFF5 uses an 8-bit length; we cap to 255 bytes.
    private static func labelCell(row: Int, column: Int, text: String) -> Data {
        var payload = Data()
        payload.append(le16(UInt16(row)))
        payload.append(le16(UInt16(column)))
        payload.append(le16(0)) // XF index

        var bytes = Array(text.utf8)
        if bytes.count > 255 { bytes = Array(bytes.prefix(255)) }
        payload.append(UInt8(bytes.count))
        payload.append(contentsOf: bytes)

        return record(sid: 0x0204, payload: payload)
    }

    // MARK: - Low-level BIFF helpers

    /// Build a BIFF record: [sid:2][length:2][payload]
    private static func record(sid: UInt16, payload: Data) -> Data {
        var out = Data()
        out.append(le16(sid))
        out.append(le16(UInt16(payload.count)))
        out.append(payload)
        return out
    }

    private static func le16(_ value: UInt16) -> Data {
        var v = value.littleEndian
        return Data(bytes: &v, count: MemoryLayout<UInt16>.size)
    }

    private static func le32(_ value: UInt32) -> Data {
        var v = value.littleEndian
        return Data(bytes: &v, count: MemoryLayout<UInt32>.size)
    }

    private static func le64(_ value: UInt64) -> Data {
        var v = value.littleEndian
        return Data(bytes: &v, count: MemoryLayout<UInt64>.size)
    }
}