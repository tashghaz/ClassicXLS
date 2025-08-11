
import Foundation

// Safe little-endian reads (no alignment assumptions; avoids arm64 traps)
enum LEb {
    @inline(__always) static func u16(_ d: Data, _ o: Int) -> UInt16 {
        precondition(o + 2 <= d.count)
        let b0 = UInt16(d[o])
        let b1 = UInt16(d[o + 1]) << 8
        return b0 | b1
    }
    @inline(__always) static func u32(_ d: Data, _ o: Int) -> UInt32 {
        precondition(o + 4 <= d.count)
        let b0 = UInt32(d[o])
        let b1 = UInt32(d[o + 1]) << 8
        let b2 = UInt32(d[o + 2]) << 16
        let b3 = UInt32(d[o + 3]) << 24
        return b0 | b1 | b2 | b3
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
        let sid = LEb.u16(bytes, offset)
        let len = Int(LEb.u16(bytes, offset + 2))
        let start = offset
        let payloadStart = offset + 4
        let end = payloadStart + len
        guard end <= bytes.count else { return nil }
        offset = end
        return BIFFRecord(sid: sid, data: bytes[payloadStart..<end], startOffset: start)
    }

    func seek(to absoluteOffset: Int) { offset = absoluteOffset }
    var isAtEnd: Bool { offset >= bytes.count }
}
