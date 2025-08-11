import Foundation

// BIFF record IDs we need now
private enum BID: UInt16 {
    case BOF        = 0x0809
    case EOF        = 0x000A
    case SST        = 0x00FC
    case CONTINUE   = 0x003C
    case BOUNDSHEET = 0x0085
}

// Helper model: one sheet entry from BOUNDSHEET
fileprivate struct BoundSheet {
    let name: String
    let streamOffset: Int
}

// Helper container for workbook "globals"
fileprivate struct WorkbookGlobals {
    let sst: [String]
    let bounds: [BoundSheet]
}

struct XLSWorkbookParser {

    /// Parse workbook globals (SST + sheet list) and return a workbook.
    /// Step 3: we only set sheet names; grids stay empty.
    static func parse(workbookStream: Data) throws -> XLSWorkbook {
        let (globals, _) = try parseGlobals(streamBytes: workbookStream)
        let sheets: [XLSSheet] = globals.bounds.map { XLSSheet(name: $0.name, grid: [:]) }
        return XLSWorkbook(sheets: sheets)
    }

    /// Scan from the start of the BIFF stream and gather SST + BOUNDSHEETs.
    /// Returns (globals, stream) — we’ll reuse the stream in Step 4 to read sheet cells.
    private static func parseGlobals(streamBytes: Data) throws -> (WorkbookGlobals, BIFFStream) {
        let s = BIFFStream(streamBytes)
        var sst: [String] = []
        var bounds: [BoundSheet] = []

        while let rec = s.next() {
            switch BID(rawValue: rec.sid) {
            case .BOF:
                // entering a substream (globals, then later each sheet) — continue
                break
            case .SST:
                sst = try parseSST(first: rec, stream: s)
            case .BOUNDSHEET:
                if let b = parseBoundSheet(rec.data) { bounds.append(b) }
            case .EOF:
                // first EOF = end of workbook globals
                return (WorkbookGlobals(sst: sst, bounds: bounds), s)
            default:
                break
            }
        }
        return (WorkbookGlobals(sst: sst, bounds: bounds), s)
    }

    // MARK: - BOUNDSHEET

    /// BOUNDSHEET structure:
    /// [0..3] absolute offset to sheet BOF
    /// [4] state (ignored)
    /// [5] type  (ignored)
    /// [6] cch (char count)
    /// [7] flags (bit0=1 -> Unicode 16-bit; else 8-bit "compressed")
    /// [8..] name bytes
    private static func parseBoundSheet(_ d: Data) -> BoundSheet? {
        guard d.count >= 8 else { return nil }
        let off = Int(LEb.u32(d, 0))  // safe u32 read
        let cch = Int(d[6])
        let flags = d[7]
        let isUnicode = (flags & 0x01) != 0

        if isUnicode {
            let need = 8 + cch * 2
            guard need <= d.count else { return nil }
            let name = String(data: d[8..<need], encoding: .utf16LittleEndian) ?? "Sheet"
            return BoundSheet(name: name, streamOffset: off)
        } else {
            let need = 8 + cch
            guard need <= d.count else { return nil }
            let name = String(data: d[8..<need], encoding: .ascii) ?? "Sheet"
            return BoundSheet(name: name, streamOffset: off)
        }
    }

    // MARK: - SST (basic)
    /// Basic Shared String Table parser.
    /// Handles common cases; skips rich/ext runs for now (good enough for Step 3).
    private static func parseSST(first: BIFFRecord, stream: BIFFStream) throws -> [String] {
        guard first.data.count >= 8 else { return [] }
        var remainingUnique = Int(LEb.u32(first.data, 4))
        var strings: [String] = []
        var chunk = first.data
        var pos = 8

        func pullString() -> String? {
            // Need 2 bytes length + 1 byte flags
            if pos + 3 > chunk.count { return nil }
            let cch = Int(LEb.u16(chunk, pos)); pos += 2
            let flags = chunk[pos]; pos += 1
            let isUnicode = (flags & 0x01) != 0
            let bytesNeeded = isUnicode ? cch * 2 : cch
            if pos + bytesNeeded > chunk.count { return nil }
            let data = chunk[pos..<(pos + bytesNeeded)]
            pos += bytesNeeded
            return isUnicode
                ? (String(data: data, encoding: .utf16LittleEndian) ?? "")
                : (String(data: data, encoding: .ascii) ?? "")
        }

        while remainingUnique > 0 {
            if let s = pullString() {
                strings.append(s)
                remainingUnique -= 1
            } else {
                // String spills into next CONTINUE
                guard let cont = stream.next(), cont.sid == BID.CONTINUE.rawValue else { break }
                chunk = cont.data
                pos = 0
            }
        }
        return strings
    }
}
