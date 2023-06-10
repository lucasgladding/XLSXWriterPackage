//
//  URL+String.swift
//  
//
//  Created by Lucas Gladding on 2023-06-08.
//

import Foundation

extension URL {
    var path: String {
        String(absoluteString.dropFirst(6))
    }
}
