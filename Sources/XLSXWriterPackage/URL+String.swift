//
//  URL+String.swift
//  
//
//  Created by Lucas Gladding on 2023-06-08.
//

import Foundation

extension URL {
    public var protocolRelativeString: String {
        let count = scheme?.count ?? 0
        return String(absoluteString.dropFirst(count + 2))
    }
}
