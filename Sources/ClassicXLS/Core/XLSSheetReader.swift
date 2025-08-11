import Foundation

// Uses the same LEb helpers defined in BIFFStream.swift

/// Minimal sheet reader: NUMBER, RK, LABELSST, LABEL, FORMULA(cached numeric)
struct XLSSheetReader {

    static func readSheet(bytes: Data, offset: Int, sst: [String]) -> [Int: [Int: XLSCell]] {
        let s = BIFFStream(bytes)
        s.seek(to: offset)

        var grid: [Int: [Int: XLSCell]] = [:]

        func put(_ row: Int, _ col: Int, _ value: XLSValue) {
            let cell = XLSCell(row: row, col: col, value: value)
            var rowDict = grid[row] ?? [:]
            rowDict[col] = cell
            grid[row] = rowDict
        }

        while let rec = s.next() {
            // Stop when the sheet ends
            if rec.sid == 0x000A /* EOF */ { break }

            switch rec.sid {
            case 0x0203: // NUMBER
                guard
                    let row = LEb.u16(rec.data, 0).map(Int.init),
                    let col = LEb.u16(rec.data, 2).map(Int.init)
                else { continue }
                // 8‐byte IEEE 754 little endian at offset 6
                if let d = readDouble(rec.data, 6) {
                    put(row, col, .number(d))
                }

            case 0x027E: // RK (compressed number)
                guard
                    let row = LEb.u16(rec.data, 0).map(Int.init),
                    let col = LEb.u16(rec.data, 2).map(Int.init),
                    let rk  = LEb.u32(rec.data, 6)
                else { continue }
                let d = decodeRK(rk)
                put(row, col, .number(d))

            case 0x00FD: // LABELSST (shared string)
                guard
                    let row = LEb.u16(rec.data, 0).map(Int.init),
                    let col = LEb.u16(rec.data, 2).map(Int.init),
                    let sstIndexU = LEb.u32(rec.data, 6)
                else { continue }
                let idx = Int(sstIndexU)
                if idx >= 0 && idx < sst.count {
                    put(row, col, .text(sst[idx]))
                }

            case 0x0204: // LABEL (old-style inline)
                guard
                    let row = LEb.u16(rec.data, 0).map(Int.init),
                    let col = LEb.u16(rec.data, 2).map(Int.init),
                    let lenU = LEb.u16(rec.data, 6)
                else { continue }
                let len = Int(lenU)
                let s = rec.data.startIndex + 8
                let e = min(s + len, rec.data.endIndex)
                if s <= e {
                    let str = String(data: rec.data[s..<e], encoding: .ascii) ?? ""
                    put(row, col, .text(str))
                }

            case 0x0006: // FORMULA (use cached numeric result if present)
                guard
                    let row = LEb.u16(rec.data, 0).map(Int.init),
                    let col = LEb.u16(rec.data, 2).map(Int.init)
                else { continue }
                // Cached result is 8 bytes at [6..14]
                if let d = readDouble(rec.data, 6), !d.isNaN {
                    put(row, col, .number(d))
                }
                // NOTE: formulas that cache strings/booleans/errors need extra handling; skipped in v0

            case 0x00BD: // MULRK (multiple RK values)
                // [0..1] row, [2..3] firstCol, then (rkrec: [xf(2) rk(4)])*, then lastCol(2)
                guard
                    let rowU = LEb.u16(rec.data, 0),
                    let firstColU = LEb.u16(rec.data, 2)
                else { continue }
                let row = Int(rowU)
                var col = Int(firstColU)
                var pos = 4
                // last 2 bytes are lastCol
                while pos + 6 <= rec.data.count - 2 {
                    // skip XF
                    pos += 2
                    guard let rk = LEb.u32(rec.data, pos) else { break }
                    pos += 4
                    put(row, col, .number(decodeRK(rk)))
                    col += 1
                }

            default:
                // ignore many other record types for MVP
                break
            }
        }

        return grid
    }

    // MARK: helpers

    /// Read a little-endian IEEE-754 Double from `data` at `offset` safely.
    private static func readDouble(_ data: Data, _ o: Int) -> Double? {
        let s = data.startIndex + o
        let e = s + 8
        guard s >= data.startIndex, e <= data.endIndex else { return nil }
        var tmp = [UInt8](repeating: 0, count: 8)
        data.copyBytes(to: &tmp, from: s..<e)
        // Compose UInt64 little-endian
        var bits: UInt64 = 0
        for i in 0..<8 {
            bits |= UInt64(tmp[i]) << (8 * i)
        }
        return Double(bitPattern: bits)
    }

    /// Decode Excel RK packed number to Double.
    /// bit0: multiply/divide by 100
    /// bit1: 1 = integer (bits 2..31 signed), 0 = floating (upper 30 bits of IEEE double)
    private static func decodeRK(_ rk: UInt32) -> Double {
        let mult100 = (rk & 0x1) != 0
        let isInt   = (rk & 0x2) != 0
        var val: Double
        if isInt {
            // signed 30-bit integer
            let v = Int32(bitPattern: rk & ~0x3) >> 2
            val = Double(v)
        } else {
            // reconstruct double: place 30 bits as the high 30 of a 64-bit value
            // (standard trick used by libxls and others)
            let high = UInt64(rk & ~0x3) << 32
            val = Double(bitPattern: high)
        }
        return mult100 ? val / 100.0 : val
    }
}
