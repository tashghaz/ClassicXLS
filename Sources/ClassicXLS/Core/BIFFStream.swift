
import Foundation

// Little-endian helpers just for this file
private enum LEb {
    static func u16(_ d: Data, _ o: Int) -> UInt16 { d.withUnsafeBytes { $0.load(fromByteOffset: o, as: UInt16.self) } }
    static func u32(_ d: Data, _ o: Int) -> UInt32 { d.withUnsafeBytes { $0.load(fromByteOffset: o, as: UInt32.self) } }
}

struct BIFFRecord {
    let sid: UInt16         // record id
    let data: Data          // payload (len bytes)
    let startOffset: Int    // absolute stream offset of this record header
}

final class BIFFStream {
    private let bytes: Data
    private(set) var offset: Int = 0
    init(_ d: Data) { self.bytes = d }

    func next() -> BIFFRecord? {
        guard offset + 4 <= bytes.count else { return nil }
        let sid = UInt16(littleEndian: LEb.u16(bytes, offset))
        let len = Int(UInt16(littleEndian: LEb.u16(bytes, offset + 2)))
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
