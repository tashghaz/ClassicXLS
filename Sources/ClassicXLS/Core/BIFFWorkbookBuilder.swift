import Foundation

/// Builds a single-sheet BIFF5 **workbook** stream (needs OLE/CFB to be a .xls).
enum BIFFWorkbookBuilder {

    static func buildWorkbook(sheetName: String, worksheetData: Data) -> Data {
        var book = Data()

        // BOF (globals)
        book.append(recBOF(.globals))

        // Essentials for BIFF5:
        book.append(recCodepage(0x04E4))         // Windows-1252
        book.append(recWindow1())
        book.append(recFont(name: "Arial", sizePt: 10))
        book.append(recXF())

        // BOUNDSHEET (offset patched later)
        let bsPos = book.count
        book.append(recBoundsheet(name: sheetName, bofOffset: 0))

        // EOF (end of globals)
        book.append(recEOF())

        // Worksheet starts here
        let wsOffset = UInt32(book.count)
        book.append(worksheetData)

        // Patch BOUNDSHEET offset
        book.replaceSubrange(bsPos + 4 ..< bsPos + 8, with: le32(wsOffset))

        return book
    }

    // MARK: records

    private enum BOFType: UInt16 { case globals = 0x0005, worksheet = 0x0010 }

    private static func recBOF(_ t: BOFType) -> Data {
        var p = Data()
        p.append(le16(0x0500))               // BIFF5
        p.append(le16(t.rawValue))
        p.append(le16(0)); p.append(le16(0)) // build id/year
        p.append(le32(0)); p.append(le32(0)) // history/lowest ver
        return rec(0x0809, p)
    }

    private static func recCodepage(_ cp: UInt16) -> Data {
        rec(0x0042, le16(cp))
    }

    private static func recWindow1() -> Data {
        var p = Data()
        p.append(le16(0))            // x
        p.append(le16(0))            // y
        p.append(le16(20000))        // width
        p.append(le16(10000))        // height
        p.append(le16(0x0038))       // flags
        p.append(le16(0))            // active tab
        p.append(le16(0))            // first visible tab
        p.append(le16(1))            // selected tabs
        p.append(le16(600))          // tab ratio
        return rec(0x003D, p)
    }

    private static func recFont(name: String, sizePt: Int) -> Data {
        var p = Data()
        p.append(le16(UInt16(sizePt * 20))) // twips
        p.append(le16(0))                   // options
        p.append(le16(0x7FFF))              // color auto
        p.append(le16(400))                 // bold
        p.append(le16(0))                   // script
        p.append(UInt8(0))                  // underline
        p.append(UInt8(0))                  // family
        p.append(UInt8(0))                  // charset
        p.append(UInt8(0))                  // reserved
        let nameBytes = Array(name.utf8.prefix(255))
        p.append(UInt8(nameBytes.count))
        p.append(contentsOf: nameBytes)
        return rec(0x0031, p)
    }

    private static func recXF() -> Data {
        var p = Data()
        p.append(le16(0))    // font index
        p.append(le16(0))    // number format
        p.append(le16(0))    // alignment/prot
        p.append(UInt8(0)); p.append(UInt8(0)); p.append(UInt8(0xF8))
        p.append(le16(0)); p.append(le16(0)); p.append(le16(0))
        return rec(0x00E0, p)
    }

    private static func recBoundsheet(name: String, bofOffset: UInt32) -> Data {
        var p = Data()
        p.append(le32(bofOffset))
        p.append(UInt8(0))   // visible
        p.append(UInt8(0))   // sheet type: worksheet
        let bytes = Array(name.utf8.prefix(31)) // 8-bit (BIFF5)
        p.append(UInt8(bytes.count))
        p.append(contentsOf: bytes)
        return rec(0x0085, p)
    }

    private static func recEOF() -> Data { rec(0x000A, Data()) }

    // MARK: helpers
    private static func rec(_ sid: UInt16, _ payload: Data) -> Data {
        var d = Data()
        d.append(le16(sid))
        d.append(le16(UInt16(payload.count)))
        d.append(payload)
        return d
    }
    private static func le16(_ v: UInt16) -> Data { var x = v.littleEndian; return Data(bytes: &x, count: 2) }
    private static func le32(_ v: UInt32) -> Data { var x = v.littleEndian; return Data(bytes: &x, count: 4) }
}