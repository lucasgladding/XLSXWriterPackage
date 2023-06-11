//
//  Range.swift
//  
//
//  Created by Lucas Gladding on 2023-06-08.
//

import Foundation
import libxlsxwriter

public struct Range {
    let location1: Location
    let location2: Location

    public init(_ row1: Int, _ col1: Int, _ row2: Int, _ col2: Int) {
        self.location1 = Location(row1, col1)
        self.location2 = Location(row2, col2)
    }

    public init(_ string: String) {
        self.init(
            Int(lxw_name_to_row(string.cString)),
            Int(lxw_name_to_col(string.cString)),
            Int(lxw_name_to_row_2(string.cString)),
            Int(lxw_name_to_col_2(string.cString))
        )
    }
}
