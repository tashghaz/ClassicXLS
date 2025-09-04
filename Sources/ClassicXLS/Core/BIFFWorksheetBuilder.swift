import Foundation

/// Builds a single BIFF5 worksheet stream (BOF → DIMENSIONS → cells → EOF).
/// NOTE: This is just the worksheet stream. It must be wrapped in a BIFF5
/// workbook stream, then into an OLE/CFB container to become a real .xls.
enum BIFFWorksheetBuilder {

    /// Headers go on row 0, columns **A..**; data starts at row 1.
    /// (Starting at A fixes the “ÿÿÿ…” garbage some apps showed when colMin ≠ 0.)
    static func makeWorksheetStream(sheetName: String,
                                    headers: [String],
                                    rows: [[String]]) -> Data {
        var s = Data()
        s.append(recBOF(.worksheet))
        s.append(recDimensions(headerCount: headers.count, rowCount: rows.count))
        s.append(recHeaderRow(headers))
        s.append(recDataRows(rows))
        s.append(recEOF())
        return s
    }

    // MARK: records

    private enum BOFType: UInt16 { case worksheet = 0x0010 }

    private static func recBOF(_ t: BOFType) -> Data {
        var p = Data()
        p.append(le16(0x0500))         // BIFF5
        p.append(le16(t.rawValue))     // worksheet
        p.append(le16(0))              // build id
        p.append(le16(0))              // build year
        p.append(le32(0))              // history flags
        p.append(le32(0))              // lowest Excel ver
        return rec(0x0809, p)
    }

    /// DIMENSIONS: rowMin,rowMax(exclusive), colMin,colMax(exclusive), flags
    private static func recDimensions(headerCount: Int, rowCount: Int) -> Data {
        let rowMin: UInt32 = 0
        let rowMax: UInt32 = UInt32(1 + rowCount) // header + data rows
        let colMin: UInt16 = 0                    // start at **A** (fixes garbled top row)
        let colMax: UInt16 = UInt16(max(headerCount, rowCount > 0 ? (rowsMaxWidth(rowCount) ?? headerCount) : headerCount))

        var p = Data()
        p.append(le32(rowMin))
        p.append(le32(rowMax))
        p.append(le16(colMin))
        p.append(le16(colMin + colMax)) // exclusive upper bound
        p.append(le16(0))
        return rec(0x0200, p)
    }

    private static func rowsMaxWidth(_ rc: Int) -> Int? { rc > 0 ? nil : nil } // keep simple

    private static func recHeaderRow(_ headers: [String]) -> Data {
        var out = Data()
        let r = 0
        for (c, text) in headers.enumerated() {
            out.append(recLabel(row: r, col: c, text: text))
        }
        return out
    }

    private static func recDataRows(_ rows: [[String]]) -> Data {
        var out = Data()
        for (ri, row) in rows.enumerated() {
            let r = 1 + ri
            for (c, raw) in row.enumerated() {
                if let d = Double(raw.replacingOccurrences(of: ",", with: ".")) {
                    out.append(recNumber(row: r, col: c, value: d))
                } else {
                    out.append(recLabel(row: r, col: c, text: raw))
                }
            }
        }
        return out
    }

    private static func recNumber(row: Int, col: Int, value: Double) -> Data {
        var p = Data()
        p.append(le16(UInt16(row)))
        p.append(le16(UInt16(col)))
        p.append(le16(0))                    // XF index
        p.append(le64(value.bitPattern))     // raw IEEE754
        return rec(0x0203, p)                // NUMBER
    }

    /// LABEL: 8-bit length + bytes encoded with Windows-1252 (BIFF5 single-byte strings)
    private static func recLabel(row: Int, col: Int, text: String) -> Data {
        var p = Data()
        p.append(le16(UInt16(row)))
        p.append(le16(UInt16(col)))
        p.append(le16(0))                    // XF index
        let bytes = dataCP1252(text, max: 255)
        p.append(UInt8(bytes.count))
        p.append(contentsOf: bytes)
        return rec(0x0204, p)                // LABEL
    }

    private static func recEOF() -> Data { rec(0x000A, Data()) }

    // MARK: helpers

    private static func dataCP1252(_ s: String, max: Int) -> [UInt8] {
        if let d = s.data(using: .windowsCP1252, allowLossyConversion: true) {
            return Array(d.prefix(max))
        }
        // fallback: replace unsupported chars with "?"
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