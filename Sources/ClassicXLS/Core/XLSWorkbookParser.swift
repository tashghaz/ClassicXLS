import Foundation

private enum BID: UInt16 {
    case BOF        = 0x0809
    case EOF        = 0x000A
    case SST        = 0x00FC
    case CONTINUE   = 0x003C
    case BOUNDSHEET = 0x0085
}

fileprivate struct BoundSheet { let name: String; let streamOffset: Int }
fileprivate struct WorkbookGlobals { let sst: [String]; let bounds: [BoundSheet] }

struct XLSWorkbookParser {

    static func parse(workbookStream: Data) throws -> XLSWorkbook {
        let (globals, _) = try parseGlobals(streamBytes: workbookStream)

        // NEW: read each sheet and build the sparse grid
        var sheets: [XLSSheet] = []
        for b in globals.bounds {
            let grid = XLSSheetReader.readSheet(bytes: workbookStream, offset: b.streamOffset, sst: globals.sst)
            sheets.append(XLSSheet(name: b.name, grid: grid))
        }
        return XLSWorkbook(sheets: sheets)
    }

    // === Globals code unchanged (safe + index-based) ===

    private static func parseGlobals(streamBytes: Data) throws -> (WorkbookGlobals, BIFFStream) {
        let s = BIFFStream(streamBytes)
        var sst: [String] = []
        var bounds: [BoundSheet] = []

        while let rec = s.next() {
            guard let id = BID(rawValue: rec.sid) else { continue }
            switch id {
            case .BOF:
                break
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

    private static func parseBoundSheet(_ d: Data) -> BoundSheet? {
        guard d.count >= 8, let off32 = LEb.u32(d, 0) else { return nil }
        let off = Int(off32)
        guard
            let cchByte = LEb.u8(d, 6),
            let flags   = LEb.u8(d, 7)
        else { return nil }
        let cch = Int(cchByte)
        let isUnicode = (flags & 0x01) != 0

        if isUnicode {
            let need = 8 + cch * 2
            guard need <= d.count else { return nil }
            let s = d.startIndex + 8, e = s + cch * 2
            let name = String(data: d[s..<e], encoding: .utf16LittleEndian) ?? "Sheet"
            return BoundSheet(name: name, streamOffset: off)
        } else {
            let need = 8 + cch
            guard need <= d.count else { return nil }
            let s = d.startIndex + 8, e = s + cch
            let name = String(data: d[s..<e], encoding: .ascii) ?? "Sheet"
            return BoundSheet(name: name, streamOffset: off)
        }
    }

    private static func parseSST(first: BIFFRecord, stream: BIFFStream) throws -> [String] {
        guard first.data.count >= 8, let unique = LEb.u32(first.data, 4) else { return [] }
        var remainingUnique = Int(unique)
        var out: [String] = []
        var chunk = first.data
        var pos = 8

        func pullString() -> String? {
            guard let cchLE = LEb.u16(chunk, pos) else { return nil }
            pos += 2
            guard let flags = LEb.u8(chunk, pos) else { return nil }
            pos += 1
            let isUnicode = (flags & 0x01) != 0
            let bytesNeeded = isUnicode ? Int(cchLE) * 2 : Int(cchLE)
            let s = chunk.startIndex + pos
            let e = s + bytesNeeded
            guard e <= chunk.endIndex else { return nil }
            let data = chunk[s..<e]
            pos += bytesNeeded
            return isUnicode
                ? (String(data: data, encoding: .utf16LittleEndian) ?? "")
                : (String(data: data, encoding: .ascii) ?? "")
        }

        while remainingUnique > 0 {
            if let s = pullString() {
                out.append(s)
                remainingUnique -= 1
            } else {
                guard let cont = stream.next(), cont.sid == BID.CONTINUE.rawValue else { break }
                chunk = cont.data
                pos = 0
            }
        }
        return out
    }
}
