//
//  File.swift
//  ClassicXLS
//
//  Created by Artashes Ghazaryan on 8/11/25.
//

import Foundation

public enum XLSValue: Equatable, CustomStringConvertible {
    case text(String)
    case number(Double)
    case date(Date)

    public var description: String {
        switch self {
        case .text(let s):   return s
        case .number(let n): return String(n)
        case .date(let d):   return ISO8601DateFormatter().string(from: d)
        }
    }
}

public struct XLSCell: Equatable {
    public let row: Int
    public let col: Int
    public let value: XLSValue
    public init(row: Int, col: Int, value: XLSValue) {
        self.row = row; self.col = col; self.value = value
    }
}

public struct XLSSheet {
    public let name: String
    /// Sparse grid: row -> (col -> cell)
    public let grid: [Int: [Int: XLSCell]]
}

public struct XLSWorkbook {
    public let sheets: [XLSSheet]
}
