//
//  FileManager+DocumentURL.swift
//  
//
//  Created by Lucas Gladding on 2023-06-08.
//

import Foundation

extension FileManager {
    func documentURL(filename: String) throws -> URL {
        try documentDirectoryURL.appendingPathComponent(filename)
    }

    private var documentDirectoryURL: URL {
        get throws {
            try url(
                for: .documentDirectory,
                in: .userDomainMask,
                appropriateFor: nil,
                create: false
            )
        }
    }
}
