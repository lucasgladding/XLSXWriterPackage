//
//  File.swift
//  
//
//  Created by Lucas Gladding on 2023-06-08.
//

import Foundation

public struct RichString {
    let string: String
    let format: Format?

    public init(_ string: String, format: Format? = nil) {
        self.string = string
        self.format = format
    }
}
