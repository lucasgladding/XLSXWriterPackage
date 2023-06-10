//
//  Document+Creation.swift
//  
//
//  Created by Lucas Gladding on 2023-06-10.
//

import Foundation

extension Document {
    convenience public init?(filename: String) {
        guard let documentURL = try? FileManager.default.documentURL(filename: filename) else {
            return nil
        }
        self.init(path: documentURL.path)
    }
}
