import Foundation

/// Single-sheet BIFF5 workbook stream (needs OLE/CFB to be a real .xls).
enum BIFFWorkbookBuilder {

    static func buildWorkbook(sheetName: String, worksheetData: Data) -> Data {
        var b = Data()
        b.append(recBOF(.globals))

        b.append(recCodepage(0x04E4))       // Windows-1252
        b.append(recWindow1())
        b.append(recFont(name: "Arial", sizePt: 10))
        b.append(recXF())

        // BOUNDSHEET (offset patched after appending sheet)
        let bsPos = b.count
        b.append(recBoundsheet(name: sheetName, bofOffset: 0))

        b.append(recEOF())

        let wsOffset = UInt32(b.count)
        b.append(worksheetData)

        // Patch absolute offset of worksheet BOF
        b.replaceSubrange(bsPos + 4 ..< bsPos + 8, with: le32(wsOffset))
        return b
    }

    private enum BOFType: UInt16 { case globals = 0x0005 }

    private static func recBOF(_ t: BOFType) -> Data {
        var p = Data()
        p.append(le16(0x0500))
        p.append(le16(t.rawValue))
        p.append(le16(0)); p.append(le16(0))
        p.append(le32(0)); p.append(le32(0))
        return rec(0x0809, p)
    }

    private static func recCodepage(_ cp: UInt16) -> Data { rec(0x0042, le16(cp)) }

    private static func recWindow1() -> Data {
        var p = Data()
        p.append(le16(0)); p.append(le16(0))
        p.append(le16(20000)); p.append(le16(10000))
        p.append(le16(0x0038))
        p.append(le16(0)); p.append(le16(0))
        p.append(le16(1)); p.append(le16(600))
        return rec(0x003D, p)
    }

    private static func recFont(name: String, sizePt: Int) -> Data {
        var p = Data()
        p.append(le16(UInt16(sizePt * 20)))
        p.append(le16(0))
        p.append(le16(0x7FFF))
        p.append(le16(400))
        p.append(le16(0))
        p.append(UInt8(0)); p.append(UInt8(0)); p.append(UInt8(0)); p.append(UInt8(0))
        let nameBytes = Array(name.utf8.prefix(255))
        p.append(UInt8(nameBytes.count))
        p.append(contentsOf: nameBytes)
        return rec(0x0031, p)
    }

    private static func recXF() -> Data {
        var p = Data()
        p.append(le16(0)); p.append(le16(0)); p.append(le16(0))
        p.append(UInt8(0)); p.append(UInt8(0)); p.append(UInt8(0xF8))
        p.append(le16(0)); p.append(le16(0)); p.append(le16(0))
        return rec(0x00E0, p)
    }

    private static func recBoundsheet(name: String, bofOffset: UInt32) -> Data {
        var p = Data()
        p.append(le32(bofOffset))
        p.append(UInt8(0))      // visible
        p.append(UInt8(0))      // worksheet
        let nb = Array(name.utf8.prefix(31))
        p.append(UInt8(nb.count))
        p.append(contentsOf: nb)
        return rec(0x0085, p)
    }

    private static func recEOF() -> Data { rec(0x000A, Data()) }

    private static func rec(_ sid: UInt16, _ payload: Data) -> Data {
        var d = Data()
        d.append(le16(sid)); d.append(le16(UInt16(payload.count))); d.append(payload)
        return d
    }
    private static func le16(_ v: UInt16) -> Data { var x = v.littleEndian; return Data(bytes: &x, count: 2) }
    private static func le32(_ v: UInt32) -> Data { var x = v.littleEndian; return Data(bytes: &x, count: 4) }
}