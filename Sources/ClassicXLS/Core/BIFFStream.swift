import Foundation

// Ultra-safe little-endian reads (no alignment assumptions, no overlapping access)
enum LEb {
    @inline(__always) static func u16(_ d: Data, _ o: Int) -> UInt16? {
        guard o >= 0, o + 2 <= d.count else { return nil }
        var tmp = [UInt8](repeating: 0, count: 2)
        d.copyBytes(to: &tmp, from: o..<(o + 2))
        return UInt16(tmp[0]) | (UInt16(tmp[1]) << 8)
    }
    @inline(__always) static func u32(_ d: Data, _ o: Int) -> UInt32? {
        guard o >= 0, o + 4 <= d.count else { return nil }
        var tmp = [UInt8](repeating: 0, count: 4)
        d.copyBytes(to: &tmp, from: o..<(o + 4))
        return UInt32(tmp[0])
             | (UInt32(tmp[1]) << 8)
             | (UInt32(tmp[2]) << 16)
             | (UInt32(tmp[3]) << 24)
    }
}

public struct BIFFRecord {
    public let sid: UInt16       // record id
    public let data: Data        // payload (len bytes)
    public let startOffset: Int  // absolute stream offset of this record header
}

final class BIFFStream {
    private let bytes: Data
    private(set) var offset: Int = 0

    init(_ d: Data) { self.bytes = d }

    func next() -> BIFFRecord? {
        guard offset + 4 <= bytes.count else { return nil }
        guard
            let sid = LEb.u16(bytes, offset),
            let lenU16 = LEb.u16(bytes, offset + 2)
        else { return nil }

        let len = Int(lenU16)
        let start = offset
        let payloadStart = offset + 4
        let end = payloadStart + len
        guard end <= bytes.count else { return nil }

        offset = end
        return BIFFRecord(sid: sid, data: bytes[payloadStart..<end], startOffset: start)
    }

    func seek(to absoluteOffset: Int) { offset = max(0, absoluteOffset) }
    var isAtEnd: Bool { offset >= bytes.count }
}
