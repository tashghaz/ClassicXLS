import Foundation

/// Builds a single BIFF5 worksheet stream (BOF → DIMENSIONS → ROWs → cells → EOF).
enum BIFFWorksheetBuilder {

    /// headers on row 0 (columns A..), data rows start at row 1
    static func makeWorksheetStream(sheetName: String,
                                    headers: [String],
                                    rows: [[String]]) -> Data {
        var ws = Data()

        ws.append(recBOF(.worksheet))

        // ---- DIMENSIONS (row/col bounds, exclusive upper limits)
        let maxRowCount = rows.count
        let maxRowWidth = rows.map { $0.count }.max() ?? 0
        let width = max(headers.count, maxRowWidth)
        ws.append(recDimensions(totalDataRows: maxRowCount, width: width))

        // ---- ROW records (header + each data row)
        ws.append(recRow(row: 0, width: headers.count))         // header row
        for (i, row) in rows.enumerated() {
            ws.append(recRow(row: 1 + i, width: row.count))
        }

        // ---- header cells (row 0, cols A..)
        for (c, text) in headers.enumerated() {
            ws.append(recLabel(row: 0, col: c, text: text))
        }

        // ---- data cells (row 1.., cols A..)
        for (ri, row) in rows.enumerated() {
            let r = 1 + ri
            for (c, raw) in row.enumerated() {
                if let d = Double(raw.replacingOccurrences(of: ",", with: ".")) {
                    ws.append(recNumber(row: r, col: c, value: d))
                } else {
                    ws.append(recLabel(row: r, col: c, text: raw))
                }
            }
        }

        ws.append(recEOF())
        return ws
    }

    // MARK: Records

    private enum BOFType: UInt16 { case worksheet = 0x0010 }

    private static func recBOF(_ type: BOFType) -> Data {
        var p = Data()
        p.append(le16(0x0500))             // BIFF5
        p.append(le16(type.rawValue))      // worksheet
        p.append(le16(0)); p.append(le16(0))
        p.append(le32(0)); p.append(le32(0))
        return rec(0x0809, p)
    }

    /// DIMENSIONS: rowMin,rowMax(exclusive), colMin,colMax(exclusive), flags
    private static func recDimensions(totalDataRows: Int, width: Int) -> Data {
        let rowMin: UInt32 = 0
        let rowMaxExclusive: UInt32 = UInt32(1 + totalDataRows)   // header + data
        let colMin: UInt16 = 0                                    // start at A
        let colMaxExclusive: UInt16 = UInt16(width)               // exclusive!

        var p = Data()
        p.append(le32(rowMin))
        p.append(le32(rowMaxExclusive))
        p.append(le16(colMin))
        p.append(le16(colMaxExclusive))
        p.append(le16(0))
        return rec(0x0200, p)
    }

    /// ROW: row, colFirst, colLast(exclusive), height, (2B), (4B), flags
    private static func recRow(row: Int, width: Int) -> Data {
        var p = Data()
        p.append(le16(UInt16(row)))
        p.append(le16(0))                          // first defined column (A)
        p.append(le16(UInt16(width)))              // last+1
        p.append(le16(0x00FF))                     // default height (approx)
        p.append(le16(0))                          // reserved
        p.append(le32(0))                          // reserved
        p.append(le16(0))                          // flags
        return rec(0x0208, p)
    }

    /// NUMBER: row, col, xf, IEEE754 double
    private static func recNumber(row: Int, col: Int, value: Double) -> Data {
        var p = Data()
        p.append(le16(UInt16(row)))
        p.append(le16(UInt16(col)))
        p.append(le16(0))                          // XF index
        p.append(le64(value.bitPattern))
        return rec(0x0203, p)
    }

    /// LABEL (BIFF5 single-byte text): row, col, xf, len(1), bytes(<=255)
    private static func recLabel(row: Int, col: Int, text: String) -> Data {
        var p = Data()
        p.append(le16(UInt16(row)))
        p.append(le16(UInt16(col)))
        p.append(le16(0))                          // XF index
        let bytes = dataCP1252(text, max: 255)
        p.append(UInt8(bytes.count))
        p.append(contentsOf: bytes)
        return rec(0x0204, p)
    }

    private static func recEOF() -> Data { rec(0x000A, Data()) }

    // MARK: Helpers

    /// Windows-1252 bytes, lossy (non-representable chars → "?")
    private static func dataCP1252(_ s: String, max: Int) -> [UInt8] {
        if let d = s.data(using: .windowsCP1252, allowLossyConversion: true) {
            return Array(d.prefix(max))
        }
        return s.map { $0.isASCII ? $0.asciiValue! : 0x3F }.prefix(max).map { $0 }
    }

    private static func rec(_ sid: UInt16, _ payload: Data) -> Data {
        var d = Data()
        d.append(le16(sid))
        d.append(le16(UInt16(payload.count)))
        d.append(payload)
        return d
    }

    private static func le16(_ v: UInt16) -> Data { var x = v.littleEndian; return Data(bytes: &x, count: 2) }
    private static func le32(_ v: UInt32) -> Data { var x = v.littleEndian; return Data(bytes: &x, count: 4) }
    private static func le64(_ v: UInt64) -> Data { var x = v.littleEndian; return Data(bytes: &x, count: 8) }
}