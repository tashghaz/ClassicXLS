import Foundation

// Little-endian helpers
private enum LE {
    static func u16(_ d: Data, _ o: Int) -> UInt16 { d.withUnsafeBytes { $0.load(fromByteOffset: o, as: UInt16.self) } }
    static func u32(_ d: Data, _ o: Int) -> UInt32 { d.withUnsafeBytes { $0.load(fromByteOffset: o, as: UInt32.self) } }
    static func i32(_ d: Data, _ o: Int) -> Int32  { d.withUnsafeBytes { $0.load(fromByteOffset: o, as: Int32.self) } }
}

/// Minimal OLE2/CFB reader: enough to extract the "Workbook"/"Book" stream for .xls
struct CFBFile {
    struct DirEntry {
        let name: String
        let type: UInt8   // 1=storage, 2=stream, 5=root
        let startSector: Int32
        let size: Int
    }

    // Fixed after init
    let data: Data
    let sectorSize: Int
    let miniSectorSize: Int
    let dirStart: Int32
    let miniFatStart: Int32
    let miniFatCount: Int32

    // Filled during init, but need to assign after some work
    var fatSectors: [Int32] = []
    var miniStreamStart: Int32 = -2
    var directory: [DirEntry] = []

    // MARK: Init

    init(fileURL: URL) throws {
        // 1) Load bytes + validate header
        self.data = try Data(contentsOf: fileURL)
        guard data.count >= 512,
              data.prefix(8) == Data([0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1]) else {
            throw XLSReadError.notXLS
        }

        // 2) Sizes (can compute immediately)
        let ssPow = Int(LE.u16(data, 30))  // sector shift
        let msPow = Int(LE.u16(data, 32))  // mini sector shift
        self.sectorSize = 1 << ssPow
        self.miniSectorSize = 1 << msPow
        guard sectorSize == 512 || sectorSize == 4096 else { throw XLSReadError.notXLS }

        // 3) Key header pointers we’ll need later
        self.dirStart     = LE.i32(data, 48)
        self.miniFatStart = LE.i32(data, 60)
        self.miniFatCount = LE.i32(data, 64)
        let difatStart    = LE.i32(data, 68)
        let difatCount    = Int(LE.u32(data, 72))

        // Local helper we can use before self is fully ready
        func sectorAt(_ sid: Int32, _ bytes: Data, _ secSize: Int) -> Data {
            let base = 512 + Int(sid) * secSize
            return bytes.subdata(in: base..<(base + secSize))
        }

        // 4) Build FAT sector list (DIFAT) without touching self methods
        var difat: [Int32] = (0..<109)
            .map { Int32(bitPattern: LE.u32(data, 76 + $0*4)) }
            .filter { $0 >= 0 }

        var cur = difatStart
        for _ in 0..<difatCount where cur >= 0 {
            let sec = sectorAt(cur, data, sectorSize)
            for i in 0..<127 {
                let v = Int32(bitPattern: LE.u32(sec, i*4))
                if v >= 0 { difat.append(v) }
            }
            cur = LE.i32(sec, 127*4) // next DIFAT sector
        }
        self.fatSectors = difat

        // 5) Read directory stream (now we can call instance methods safely)
        let dirData = try readChain(startSector: dirStart)
        let dirCount = dirData.count / 128

        var entries: [DirEntry] = []
        var rootMiniStart: Int32 = -2

        for i in 0..<dirCount {
            let off = i * 128
            let nameLen = Int(LE.u16(dirData, off + 64))  // bytes, includes trailing null
            let rawName = dirData.subdata(in: off..<(off + 64))
            let nameData = rawName.prefix(max(0, nameLen - 2)) // drop trailing null (2 bytes)
            let name = String(data: nameData, encoding: .utf16LittleEndian) ?? ""

            let type  = dirData[off + 66]
            let start = LE.i32(dirData, off + 116)
            let size  = Int(LE.u32(dirData, off + 120))

            entries.append(DirEntry(name: name, type: type, startSector: start, size: size))
            if type == 5 { rootMiniStart = start }
        }

        self.directory = entries
        self.miniStreamStart = rootMiniStart
    }

    // MARK: FAT helpers

    private func sectorData(at sid: Int32) -> Data {
        let base = 512 + Int(sid) * sectorSize
        return data.subdata(in: base..<(base + sectorSize))
    }

    private func fatNext(_ sid: Int32) -> Int32 {
        let perSec = sectorSize / 4
        let which  = Int(sid) / perSec
        let off    = (Int(sid) % perSec) * 4
        let fatSec = fatSectors[which]
        let sec    = sectorData(at: fatSec)
        return LE.i32(sec, off)
    }

    func readChain(startSector: Int32, exactSize: Int? = nil) throws -> Data {
        var sid = startSector
        var chunks: [Data] = []
        var guardCnt = 0
        while sid >= 0 {
            chunks.append(sectorData(at: sid))
            sid = fatNext(sid)
            guardCnt += 1
            if guardCnt > 1_000_000 { throw XLSReadError.parseError("FAT loop") }
        }
        var out = Data(chunks.joined())
        if let n = exactSize, out.count > n { out = out.prefix(n) }
        return out
    }

    private func miniFatNext(_ msid: Int32, miniFat: Data) -> Int32 {
        let off = Int(msid) * 4
        return LE.i32(miniFat, off)
    }

    func readMiniChain(startMini: Int32, exactSize: Int) throws -> Data {
        // Entire MiniStream lives under the root storage chain
        let bigMini = try readChain(startSector: miniStreamStart)
        // MiniFAT itself is a regular FAT chain
        let miniFat = (miniFatCount > 0)
            ? try readChain(startSector: miniFatStart, exactSize: Int(miniFatCount) * sectorSize)
            : Data()

        var msid = startMini
        var chunks: [Data] = []
        var have = 0
        var guardCnt = 0

        while msid >= 0 && have < exactSize {
            let off = Int(msid) * miniSectorSize
            chunks.append(bigMini.subdata(in: off..<(off + miniSectorSize)))
            have += miniSectorSize
            msid = miniFatNext(msid, miniFat: miniFat)
            guardCnt += 1
            if guardCnt > 1_000_000 { throw XLSReadError.parseError("MiniFAT loop") }
        }

        var out = Data(chunks.joined())
        if out.count > exactSize { out = out.prefix(exactSize) }
        return out
    }

    // MARK: Public

    func stream(named target: String) throws -> Data {
        guard let e = directory.first(where: {
            $0.type == 2 && $0.name.caseInsensitiveCompare(target) == .orderedSame
        }) else {
            throw XLSReadError.workbookStreamMissing
        }

        if e.size < 4096 {
            return try readMiniChain(startMini: e.startSector, exactSize: e.size)
        } else {
            return try readChain(startSector: e.startSector, exactSize: e.size)
        }
    }
}
