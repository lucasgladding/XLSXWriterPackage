//
//  Location.swift
//  
//
//  Created by Lucas Gladding on 2023-06-08.
//

import Foundation
import libxlsxwriter

public struct Location {
    let row: Int
    let col: Int

    public init(_ row: Int, _ col: Int) {
        self.row = row
        self.col = col
    }

    public init(_ string: String) {
        self.init(
            Int(lxw_name_to_row(string.cString)),
            Int(lxw_name_to_col(string.cString))
        )
    }
}
