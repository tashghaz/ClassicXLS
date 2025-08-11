
import Foundation

// BIFF record iterator (to be implemented in Step 3)
struct BIFFRecord {
    let sid: UInt16
    let data: Data
}

final class BIFFStream {
    private let bytes: Data
    private(set) var offset: Int = 0
    init(_ d: Data) { self.bytes = d }

    func next() -> BIFFRecord? {
        guard offset + 4 <= bytes.count else { return nil }
        let sid = UInt16(littleEndian: bytes.withUnsafeBytes { $0.load(fromByteOffset: offset, as: UInt16.self) })
        let lenLE = bytes.withUnsafeBytes { $0.load(fromByteOffset: offset+2, as: UInt16.self) }
        let len = Int(UInt16(littleEndian: lenLE))
        let start = offset + 4
        let end = start + len
        guard end <= bytes.count else { return nil }
        offset = end
        return BIFFRecord(sid: sid, data: bytes[start..<end])
    }

    func seek(to absoluteOffset: Int) { offset = absoluteOffset }
}
