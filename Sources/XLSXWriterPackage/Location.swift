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
}

extension Location {
    public init(_ string: String) {
        let components = string.withCString { cString in
            return (
                Int(lxw_name_to_row(cString)),
                Int(lxw_name_to_col(cString))
            )
        }
        self.init(components.0, components.1)
    }
}
