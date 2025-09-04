import Foundation

/// Minimal OLE/CFB writer for one regular FAT stream named `streamName`.
/// We pad the stream to >= 4096 bytes so it goes into regular FAT
/// (no MiniFAT), which keeps the writer small and Excel-compatible.
enum CFBWriter {
    private static let sectorSize = 512
    private static let ENDOFCHAIN: UInt32 = 0xFFFFFFFE
    private static let FATSECT:    UInt32 = 0xFFFFFFFD
    private static let FREESECT:   UInt32 = 0xFFFFFFFF

    /// Write a single Compound File with one stream (e.g. "Book") to disk.
    static func writeSingleStream(streamName: String, stream: Data, to url: URL) throws {
        // 1) Ensure the stream is large enough for regular FAT (>= 4096 bytes)
        var payload = stream
        if payload.count < 4096 {
            payload.append(Data(repeating: 0, count: 4096 - payload.count))
        }

        // 2) Build the Directory stream (Root + stream entry)
        let directoryStream = buildDirectory(streamName: streamName, streamSize: UInt64(payload.count))

        // 3) Split payloads into 512-byte sectors
        let payloadSectors = splitIntoSectors(payload)
        let directorySectors = splitIntoSectors(directoryStream)

        // Layout after the 512-byte file header:
        // [Payload sectors][Directory sectors][FAT sector]
        let payloadStartSID = 0
        let directoryStartSID = payloadSectors.count
        let fatStartSID = directoryStartSID + directorySectors.count

        // 4) FAT (one sector is enough for small files: 512/4 = 128 entries)
        let totalSectors = payloadSectors.count + directorySectors.count + 1 // +1 FAT
        precondition(totalSectors <= 128, "File too large for this minimal writer (needs >1 FAT sector).")

        var fatEntries = Array(repeating: FREESECT, count: totalSectors)

        // Chain payload sectors
        for i in 0..<payloadSectors.count {
            fatEntries[payloadStartSID + i] = (i == payloadSectors.count - 1)
                ? ENDOFCHAIN
                : UInt32(payloadStartSID + i + 1)
        }

        // Directory sector (single chain)
        fatEntries[directoryStartSID] = ENDOFCHAIN

        // FAT sector marks itself
        fatEntries[fatStartSID] = FATSECT

        // 5) Build file header
        let header = buildHeader(
            numberOfFATSectors: 1,
            firstDirectorySID: UInt32(directoryStartSID),
            firstMiniFATSID: ENDOFCHAIN, numberOfMiniFATSectors: 0,
            firstDIFATSID: ENDOFCHAIN,   numberOfDIFATSectors: 0,
            firstFATSIDInDIFAT: UInt32(fatStartSID)
        )

        // 6) Assemble file bytes
        var file = Data()
        file.append(header)
        payloadSectors.forEach { file.append($0) }
        directorySectors.forEach { file.append($0) }
        file.append(buildFATSector(fatEntries))

        try file.write(to: url, options: .atomic)
    }

    // MARK: - Header / FAT

    private static func buildHeader(
        numberOfFATSectors: UInt32,
        firstDirectorySID: UInt32,
        firstMiniFATSID: UInt32, numberOfMiniFATSectors: UInt32,
        firstDIFATSID: UInt32,   numberOfDIFATSectors: UInt32,
        firstFATSIDInDIFAT: UInt32
    ) -> Data {
        var d = Data(count: sectorSize)

        func put<T: FixedWidthInteger>(_ v: T, at offset: Int) {
            var x = v.littleEndian
            d.replaceSubrange(offset ..< offset + MemoryLayout<T>.size,
                              with: Data(bytes: &x, count: MemoryLayout<T>.size))
        }

        // Signature D0 CF 11 E0 A1 B1 1A E1
        d.replaceSubrange(0..<8, with: Data([0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1]))

        // Minor(0x003E), Major(0x0003 => 512-byte sectors)
        put(UInt16(0x003E), at: 24)
        put(UInt16(0x0003), at: 26)
        // Sector sizes
        put(UInt16(0x0009), at: 28) // sector shift 2^9 = 512
        put(UInt16(0x0006), at: 30) // mini sector shift 2^6 = 64

        // Directory sectors (v4 only) = 0
        put(UInt32(0), at: 40)
        // FAT count
        put(numberOfFATSectors, at: 44)
        // First directory sector id
        put(firstDirectorySID, at: 48)
        // Transaction signature
        put(UInt32(0), at: 52)
        // Mini stream cutoff (4096)
        put(UInt32(4096), at: 56)
        // MiniFAT chain
        put(firstMiniFATSID, at: 60)
        put(numberOfMiniFATSectors, at: 64)
        // DIFAT chain
        put(firstDIFATSID, at: 68)
        put(numberOfDIFATSectors, at: 72)

        // DIFAT[0..108] table at byte 76 — put our single FAT sector SID, rest FREESECT
        put(firstFATSIDInDIFAT, at: 76)
        for i in 1..<109 {
            put(FREESECT, at: 76 + i * 4)
        }
        return d
    }

    private static func buildFATSector(_ entries: [UInt32]) -> Data {
        var d = Data(capacity: sectorSize)
        for v in entries {
            var x = v.littleEndian
            d.append(Data(bytes: &x, count: 4))
        }
        if d.count < sectorSize {
            d.append(Data(repeating: 0, count: sectorSize - d.count))
        }
        return d
    }

    // MARK: - Directory (Root + one Stream)

    private static func buildDirectory(streamName: String, streamSize: UInt64) -> Data {
        var dir = Data()
        dir.append(buildDirectoryEntry(name: "Root Entry", type: 5, startSector: ENDOFCHAIN, size: 0))
        dir.append(buildDirectoryEntry(name: streamName, type: 2, startSector: 0, size: streamSize))
        if dir.count < sectorSize {
            dir.append(Data(repeating: 0, count: sectorSize - dir.count))
        }
        return dir
    }

    /// Build a 128-byte directory entry.
    /// type: 5=root, 2=stream
    private static func buildDirectoryEntry(name: String, type: UInt8, startSector: UInt32, size: UInt64) -> Data {
        var e = Data(count: 128)

        // Name: UTF-16LE + trailing null, max 32 chars incl null => 64 bytes
        let trimmed = String(name.prefix(31))
        let nameLE = (trimmed.data(using: .utf16LittleEndian) ?? Data()) + Data([0, 0])
        let nameField = nameLE.prefix(64)
        e.replaceSubrange(0..<nameField.count, with: nameField)

        // Name length (bytes incl null)
        var nameLenLE = UInt16(nameField.count).littleEndian
        e.replaceSubrange(64..<66, with: Data(bytes: &nameLenLE, count: 2))

        e[66] = type     // object type
        e[67] = 0        // color (red/black) — irrelevant here

        // left/right/child = -1
        e.replaceSubrange(68..<80, with: Data(repeating: 0xFF, count: 12))

        // CLSID(16), state bits(4), timestamps(16) — left zero

        // starting sector id
        var ss = startSector.littleEndian
        e.replaceSubrange(116..<120, with: Data(bytes: &ss, count: 4))

        // stream size (bytes)
        var sz = size.littleEndian
        e.replaceSubrange(120..<128, with: Data(bytes: &sz, count: 8))

        return e
    }

    // MARK: - Utils

    private static func splitIntoSectors(_ data: Data) -> [Data] {
        var sectors: [Data] = []
        var i = 0
        while i < data.count {
            let end = min(i + sectorSize, data.count)
            var chunk = data.subdata(in: i..<end)
            if chunk.count < sectorSize {
                chunk.append(Data(repeating: 0, count: sectorSize - chunk.count))
            }
            sectors.append(chunk)
            i = end
        }
        if sectors.isEmpty { sectors = [Data(repeating: 0, count: sectorSize)] }
        return sectors
    }
}