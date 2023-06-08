//
//  String+Custom.swift
//  
//
//  Created by Lucas Gladding on 2023-06-10.
//

import Foundation

extension String {
    var cString: UnsafeMutablePointer<CChar>? {
        var cString = (self as NSString).utf8String
        return UnsafeMutablePointer<CChar>(mutating: cString)
    }
}
