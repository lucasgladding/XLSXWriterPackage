//
//  CustomExampleTests.swift
//
//
//  Created by Lucas Gladding on 2022-12-04.
//

import XCTest
import XLSXWriterPackage

final class CustomExampleTests: XCTestCase {

    func testBorders() throws {
        let document = Document(filename: "hello_world.xlsx")!
        let sheet = document.sheet()!

        let borderFormat = document.format()!
        borderFormat.set(.top(.thin))
        borderFormat.set(.bottom(.double))

        sheet.write(.string("Hello"), location: Location(0, 0))
        sheet.write(.number(123), location: Location(1, 0), format: borderFormat)

        document.close()
    }

}
