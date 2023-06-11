//
//  File.swift
//  
//
//  Created by Lucas Gladding on 2023-06-08.
//

import Foundation
import libxlsxwriter

public struct RichString {
    let string: String
    let format: Format?

    public init(_ string: String, format: Format? = nil) {
        self.string = string
        self.format = format
    }
}

extension RichString {
    var lxwRichStringTuple: lxw_rich_string_tuple {
        lxw_rich_string_tuple(
            format: format?.lxw_format,
            string: string.cString
        )
    }
}
