import Foundation

// IDs we care about now
private enum BID: UInt16 {
    case BOF        = 0x0809
    case EOF        = 0x000A
    case SST        = 0x00FC
    case CONTINUE   = 0x003C
    case BOUNDSHEET = 0x0085
}

// BOUNDSHEET entry (helper model)
fileprivate struct BoundSheet {
    let name: String
    let streamOffset: Int
}

// Minimal workbook globals (helper container)
fileprivate struct WorkbookGlobals {
    let sst: [String]
    let bounds: [BoundSheet]
}

struct XLSWorkbookParser {

    // Entry point: parse workbook globals (sheet names/offsets + SST)
    static func parse(workbookStream: Data) throws -> XLSWorkbook {
        let (globals, _) = try parseGlobals(streamBytes: workbookStream)

        // Step 3: only fill sheet names; grids stay empty for now
        let sheets: [XLSSheet] = globals.bounds.map {
            XLSSheet(name: $0.name, grid: [:])
        }
        return XLSWorkbook(sheets: sheets)
    }

    // Parses from beginning of stream; returns globals + the BIFF stream for later use
    private static func parseGlobals(streamBytes: Data) throws -> (WorkbookGlobals, BIFFStream) {
        let s = BIFFStream(streamBytes)
        var sst: [String] = []
        var bounds: [BoundSheet] = []

        while let rec = s.next() {
            switch BID(rawValue: rec.sid) {
            case .BOF:
                break // inside a section
            case .SST:
                sst = try parseSST(first: rec, stream: s)
            case .BOUNDSHEET:
                if let b = parseBoundSheet(rec.data) { bounds.append(b) }
            case .EOF:
                return (WorkbookGlobals(sst: sst, bounds: bounds), s)
            default:
                break
            }
        }
        return (WorkbookGlobals(sst: sst, bounds: bounds), s)
    }

    // MARK: - BOUNDSHEET

    // BOUNDSHEET record:
    // [0..3] stream offset to sheet BOF
    // [4] state, [5] type (ignored here)
    // [6] cch, [7] flags (bit0 = Unicode)
    private static func parseBoundSheet(_ d: Data) -> BoundSheet? {
        guard d.count >= 8 else { return nil }
        let off = Int(UInt32(littleEndian: d.withUnsafeBytes { $0.load(fromByteOffset: 0, as: UInt32.self) }))
        let cch = Int(d[6])
        let flags = d[7]
        let isUnicode = (flags & 0x01) != 0
        let name: String
        if isUnicode {
            let need = 8 + cch * 2
            guard need <= d.count else { return nil }
            name = String(data: d[8..<need], encoding: .utf16LittleEndian) ?? "Sheet"
        } else {
            let need = 8 + cch
            guard need <= d.count else { return nil }
            name = String(data: d[8..<need], encoding: .ascii) ?? "Sheet"
        }
        return BoundSheet(name: name, streamOffset: off)
    }

    // MARK: - SST (basic). Handles common strings; skips rich/ext runs for now.
    private static func parseSST(first: BIFFRecord, stream: BIFFStream) throws -> [String] {
        guard first.data.count >= 8 else { return [] }
        var remainingUnique = Int(UInt32(littleEndian: first.data.withUnsafeBytes { $0.load(fromByteOffset: 4, as: UInt32.self) }))
        var strings: [String] = []
        var chunk = first.data
        var pos = 8

        func pullString() -> String? {
            if pos + 3 > chunk.count { return nil }
            let cch = Int(UInt16(littleEndian: chunk.withUnsafeBytes { $0.load(fromByteOffset: pos, as: UInt16.self) })); pos += 2
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
                guard let cont = stream.next(), cont.sid == BID.CONTINUE.rawValue else { break }
                chunk = cont.data
                pos = 0
            }
        }
        return strings
    }
}
