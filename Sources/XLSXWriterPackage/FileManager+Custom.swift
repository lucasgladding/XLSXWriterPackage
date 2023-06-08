//
//  FileManager+Custom.swift
//  
//
//  Created by Lucas Gladding on 2023-06-08.
//

import Foundation

extension FileManager {
    func documentURL(filename: String) throws -> URL {
        let documentURL = try! FileManager.default.url(
            for: .documentDirectory,
            in: .userDomainMask,
            appropriateFor: nil,
            create: false
        )
        return documentURL.appendingPathComponent(filename)
    }
}
