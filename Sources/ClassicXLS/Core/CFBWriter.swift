import Foundation

/// Minimal OLE/CFB writer for a *single* stream (e.g. "Book").
/// - 512-byte sectors (ver 3)
/// - No MiniFAT (we pad the stream to ≥4096 and sector-align)
/// - Directory: Root Entry + your stream
enum CFBWriter {

    /// Old API your code calls.
    static func writeSingleStream(streamName: String = "Book",
                                  stream: Data,
                                  to url: URL) throws {
        let file = makeOLEFile(streamName: streamName, stream: stream)
        try file.write(to: url)
    }

    // MARK: builder

    static func makeOLEFile(streamName: String = "Book", stream: Data) -> Data {
        let sector = 512
        let entriesPerFATSector = sector / 4

        // 1) Pad main stream to ≥4096 and sector-align → no MiniFAT.
        var main = stream
        if main.count < 4096 { main.append(Data(repeating: 0, count: 4096 - main.count)) }
        pad(&main, to: sector)

        // 2) Directory stream (Root + stream)
        let mainFirstSector = 0
        let dir = makeDirectory(streamName: streamName,
                                mainStart: Int32(mainFirstSector),
                                mainSize: UInt64(main.count),
                                sector: sector)

        let mainSectors = main.count / sector
        let dirSectors  = dir.count  / sector

        // 3) FAT sectors required
        var fatSectors = 1
        while (mainSectors + dirSectors + fatSectors) > fatSectors * entriesPerFATSector {
            fatSectors += 1
        }

        let dirFirstSector = mainFirstSector + mainSectors
        let fatFirstSector = dirFirstSector + dirSectors
        let totalSectors   = mainSectors + dirSectors + fatSectors

        // 4) FAT table
        var fat = [UInt32](repeating: FREE, count: totalSectors)

        // chain main
        for i in 0..<mainSectors {
            let sid = mainFirstSector + i
            fat[sid] = (i == mainSectors - 1) ? EOC : UInt32(sid + 1)
        }
        // chain directory
        for i in 0..<dirSectors {
            let sid = dirFirstSector + i
            fat[sid] = (i == dirSectors - 1) ? EOC : UInt32(sid + 1)
        }
        // mark FAT sectors
        for i in 0..<fatSectors { fat[fatFirstSector + i] = FATSECT }

        // serialize FAT
        var fatData = Data(capacity: fatSectors * sector)
        for v in fat {
            var le = v.littleEndian
            withUnsafeBytes(of: &le) { fatData.append(contentsOf: $0) }
        }
        pad(&fatData, to: sector)

        // 5) header
        let hdr = makeHeader(firstDirSector: Int32(dirFirstSector),
                             fatSectorCount: fatSectors,
                             fatSectorIndices: (0..<fatSectors).map { UInt32(fatFirstSector + $0) })

        // 6) assemble
        var file = Data()
        file.append(hdr)
        file.append(main)
        file.append(dir)
        file.append(fatData)
        return file
    }

    // MARK: directory (2 entries)

    private static func makeDirectory(streamName: String,
                                      mainStart: Int32,
                                      mainSize: UInt64,
                                      sector: Int) -> Data {
        var d = Data(capacity: sector)

        // entry 0: Root
        d.append(dirEntry(name: "Root Entry", type: 5,
                          left: NOSTREAM, right: NOSTREAM, child: 1,
                          start: NOSTREAM, size: 0))

        // entry 1: Your stream (e.g. "Book")
        d.append(dirEntry(name: streamName, type: 2,
                          left: NOSTREAM, right: NOSTREAM, child: NOSTREAM,
                          start: mainStart, size: mainSize))

        // pad to sectors
        while d.count % sector != 0 {
            d.append(Data(repeating: 0, count: 128))
        }
        return d
    }

    private static func dirEntry(name: String, type: UInt8,
                                 left: Int32, right: Int32, child: Int32,
                                 start: Int32, size: UInt64) -> Data {
        var e = Data(capacity: 128)

        let (nameBuf, lenLE) = utf16LEName(name, maxChars: 64)
        e.append(nameBuf)            // 128-byte name buffer
        e.append(lenLE)              // length in bytes incl. null

        e.append(type)               // 1=storage 2=stream 5=root
        e.append(UInt8(0))           // color
        e.append(int32LE(left))
        e.append(int32LE(right))
        e.append(int32LE(child))
        e.append(Data(repeating: 0, count: 16)) // clsid
        e.append(uint32LE(0))        // state bits
        e.append(uint64LE(0))        // create time
        e.append(uint64LE(0))        // mod time
        e.append(int32LE(start))
        e.append(uint64LE(size))
        return e
    }

    private static func utf16LEName(_ name: String, maxChars: Int) -> (Data, Data) {
        let chars = Array(name.utf16.prefix(maxChars - 1))
        var buf = Data(count: maxChars * 2)
        for (i, u) in chars.enumerated() {
            var le = u.littleEndian
            withUnsafeBytes(of: &le) { buf.replaceSubrange(i*2..<(i*2+2), with: $0) }
        }
        let lenBytes = UInt16((chars.count + 1) * 2).littleEndian
        return (buf, Data(bytes: UnsafePointer([lenBytes]), count: 2))
    }

    // MARK: header (512 bytes)

    private static func makeHeader(firstDirSector: Int32,
                                   fatSectorCount: Int,
                                   fatSectorIndices: [UInt32]) -> Data {
        var h = Data(count: 512)
        var cur = 0
        func put(_ d: Data) { h.replaceSubrange(cur..<cur+d.count, with: d); cur += d.count }

        put(Data([0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1]))  // signature
        put(Data(repeating: 0, count: 16))                    // clsid
        put(uint16LE(0x003E))                                 // minor
        put(uint16LE(0x0003))                                 // major (512B sectors)
        put(uint16LE(0xFFFE))                                 // byte order
        put(uint16LE(9))                                      // sectorShift=9 → 512
        put(uint16LE(6))                                      // miniSectorShift=6 → 64
        put(Data(repeating: 0, count: 6))
        put(uint32LE(0))                                      // dir sector count (unused)
        put(uint32LE(UInt32(fatSectorCount)))
        put(int32LE(firstDirSector))
        put(uint32LE(0))                                      // transaction sig
        put(uint32LE(4096))                                   // mini cutoff
        put(int32LE(NOSTREAM)); put(uint32LE(0))              // miniFAT
        put(int32LE(NOSTREAM)); put(uint32LE(0))              // DIFAT chain

        // DIFAT[109]
        var difat = fatSectorIndices + Array(repeating: 0xFFFFFFFF, count: max(0, 109 - fatSectorIndices.count))
        for i in 0..<109 { put(uint32LE(difat[i])) }
        return h
    }

    // MARK: utils

    private static func pad(_ d: inout Data, to size: Int) {
        let r = d.count % size
        if r != 0 { d.append(Data(repeating: 0, count: size - r)) }
    }
    private static func uint16LE(_ v: UInt16) -> Data { var x = v.littleEndian; return Data(bytes: &x, count: 2) }
    private static func uint32LE(_ v: UInt32) -> Data { var x = v.littleEndian; return Data(bytes: &x, count: 4) }
    private static func uint64LE(_ v: UInt64) -> Data { var x = v.littleEndian; return Data(bytes: &x, count: 8) }
    private static func int32LE(_ v: Int32) -> Data { var x = v.littleEndian; return Data(bytes: &x, count: 4) }

    private static let FREE:    UInt32 = 0xFFFFFFFF
    private static let EOC:     UInt32 = 0xFFFFFFFE
    private static let FATSECT: UInt32 = 0xFFFFFFFD
    private static let NOSTREAM: Int32 = -1
}