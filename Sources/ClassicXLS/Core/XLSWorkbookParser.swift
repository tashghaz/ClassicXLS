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
        let sheets = globals.bounds.map { XLSSheet(name: $0.name, grid: [:]) }
        return XLSWorkbook(sheets: sheets)
    }

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

    // MARK: - BOUNDSHEET

    private static func parseBoundSheet(_ d: Data) -> BoundSheet? {
        guard d.count >= 8, let off32 = LEb.u32(d, 0) else { return nil }
        let off = Int(off32)
        guard d.count >= 8 else { return nil }
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
    private static func parseSST(first: BIFFRecord, stream: BIFFStream) throws -> [String] {
        guard first.data.count >= 8, let unique = LEb.u32(first.data, 4) else { return [] }
        var remainingUnique = Int(unique)
        var out: [String] = []
        var chunk = first.data
        var pos = 8

        func pullString() -> String? {
            guard let cchLE = LEb.u16(chunk, pos) else { return nil }
            pos += 2
            guard pos < chunk.count else { return nil }
            let flags = chunk[pos]; pos += 1
            let isUnicode = (flags & 0x01) != 0
            let bytesNeeded = isUnicode ? Int(cchLE) * 2 : Int(cchLE)
            guard pos + bytesNeeded <= chunk.count else { return nil }
            let data = chunk[pos..<(pos + bytesNeeded)]
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
