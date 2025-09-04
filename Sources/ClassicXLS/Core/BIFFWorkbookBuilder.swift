import Foundation

/// Builds a BIFF5 **workbook** stream that wraps one worksheet stream.
/// Layout:
///   BOF(WorkbookGlobals) → BOUNDSHEET (with offset) → EOF → [WorksheetStream bytes]
enum BIFFWorkbookBuilder {

    /// Build the workbook (globals) stream around a single worksheet stream.
    /// - Parameters:
    ///   - sheetName: visible name of the worksheet tab (<=31 bytes ASCII recommended for BIFF5)
    ///   - worksheetStream: the BIFF5 worksheet bytes (from Step 1)
    /// - Returns: BIFF5 workbook stream bytes (`Data`)
    static func makeWorkbookStream(sheetName: String, worksheetStream: Data) -> Data {
        var workbookData = Data()

        // 1) BOF: Workbook Globals (sid 0x0809, BIFF5)
        var bofPayload = Data()
        bofPayload.append(le16(0x0500))  // BIFF5 version
        bofPayload.append(le16(0x0005))  // BOF type: Workbook Globals
        bofPayload.append(le16(0))       // build id
        bofPayload.append(le16(0))       // build year
        bofPayload.append(le32(0))       // file history flags
        bofPayload.append(le32(0))       // lowest Excel version
        workbookData.append(record(sid: 0x0809, payload: bofPayload))

        // 2) BOUNDSHEET (sid 0x0085) — will patch the sheet offset after appending
        let nameBytes = Data(sheetName.utf8.prefix(31))   // BIFF5 uses 8-bit length here
        var boundsheetPayload = Data()
        boundsheetPayload.append(le32(0))                 // placeholder: absolute offset of sheet BOF
        boundsheetPayload.append(0x00)                    // visibility (0 = visible)
        boundsheetPayload.append(0x00)                    // sheet type (0x00 = worksheet)
        boundsheetPayload.append(UInt8(nameBytes.count))  // name length (8-bit)
        boundsheetPayload.append(nameBytes)

        let boundsheetRecordStart = workbookData.count
        workbookData.append(record(sid: 0x0085, payload: boundsheetPayload))

        // 3) EOF: end of workbook globals
        workbookData.append(record(sid: 0x000A, payload: Data()))

        // 4) Append the worksheet stream
        let sheetBOFAbsoluteOffset = UInt32(workbookData.count)
        workbookData.append(worksheetStream)

        // 5) Patch the BOUNDSHEET offset (first 4 bytes of its payload)
        //    Record header is 4 bytes → payload begins at boundsheetRecordStart + 4
        let offsetField = (boundsheetRecordStart + 4) ..< (boundsheetRecordStart + 8)
        workbookData.replaceSubrange(offsetField, with: le32(sheetBOFAbsoluteOffset))

        return workbookData
    }

    // MARK: - Low-level helpers (same style as Step 1)
    private static func record(sid: UInt16, payload: Data) -> Data {
        var d = Data()
        d.append(le16(sid))
        d.append(le16(UInt16(payload.count)))
        d.append(payload)
        return d
    }

    private static func le16(_ value: UInt16) -> Data {
        var v = value.littleEndian
        return Data(bytes: &v, count: MemoryLayout<UInt16>.size)
    }

    private static func le32(_ value: UInt32) -> Data {
        var v = value.littleEndian
        return Data(bytes: &v, count: MemoryLayout<UInt32>.size)
    }
}