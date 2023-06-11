//
//  URLTests.swift
//  
//
//  Created by Lucas Gladding on 2023-06-10.
//

import XCTest


final class URLTests: XCTestCase {

    func testProtocolRelativeString() throws {
        let url = URL(string: "file://Users/lucas/Documents/document.xlsx")!
        XCTAssertEqual(url.protocolRelativeString, "/Users/lucas/Documents/document.xlsx")
    }

}
