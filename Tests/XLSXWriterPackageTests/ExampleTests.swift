//
//  ExampleTests.swift
//
//
//  Created by Lucas Gladding on 2022-12-04.
//

import XCTest
import XLSXWriterPackage

final class ExampleTests: XCTestCase {

    func testHelloWorld() throws {
        let document = Document(filename: "hello_world.xlsx")!
        let sheet = document.sheet()!

        sheet.write(.string("Hello"), location: Location(0, 0))
        sheet.write(.number(123), location: Location(1, 0))

        document.close()
    }

    func testAnatomy() throws {
        let document = Document(filename: "anatomy.xlsx")!
        let sheet1 = document.sheet(name: "Demo")!
        let sheet2 = document.sheet()!

        let format1 = document.format()!
        let format2 = document.format()!

        format1.set(.bold)
        format2.set(.number("$#,##0.00"))

        sheet1.column(width: 20, first: 0, last: 0)

        sheet1.write(.string("Peach"), location: Location(0, 0))
        sheet1.write(.string("Plum"), location: Location(1, 0))
        sheet1.write(.string("Pear"), location: Location(2, 0), format: format1)
        sheet1.write(.string("Persimmon"), location: Location(3, 0), format: format1)
        sheet1.write(.number(123), location: Location(5, 0))
        sheet1.write(.number(4567.555), location: Location(6, 0), format: format2)

        sheet2.write(.string("Some text"), location: Location(0, 0), format: format1)

        document.close()
    }

    func testDemo() throws {
        let document = Document(filename: "demo.xlsx")!
        let sheet = document.sheet()!

        let format = document.format()!

        format.set(.bold)

        sheet.column(width: 20, first: 0, last: 0)

        sheet.write(.string("Hello"), location: Location(0, 0))
        sheet.write(.string("World"), location: Location(1, 0), format: format)
        sheet.write(.number(123), location: Location(2, 0))
        sheet.write(.number(123.456), location: Location(3, 0))
        sheet.write(.image("logo.png"), location: Location(1, 2))

        document.close()
    }

    func testTutorial01() throws {
        struct Expense {
            let name: String
            let cost: Double
        }

        let expenses: [Expense] = [
            Expense(name: "Rent", cost: 1000),
            Expense(name: "Gas", cost: 100),
            Expense(name: "Food", cost: 300),
            Expense(name: "Gym", cost: 50)
        ]

        let document = Document(filename: "tutorial01.xlsx")!
        let sheet = document.sheet()!

        var row: Int = 0
        let col: Int = 0

        for expense in expenses {
            sheet.write(.string(expense.name), location: Location(row, col))
            sheet.write(.number(expense.cost), location: Location(row, col + 1))
            row += 1
        }

        sheet.write(.string("Total"), location: Location(row, col))
        sheet.write(.formula("=SUM(B1:B4)"), location: Location(row, col + 1))

        document.close()
    }

    func testTutorial02() throws {
        struct Expense {
            let name: String
            let cost: Double
        }

        let expenses: [Expense] = [
            Expense(name: "Rent", cost: 1000),
            Expense(name: "Gas", cost: 100),
            Expense(name: "Food", cost: 300),
            Expense(name: "Gym", cost: 50)
        ]

        let document = Document(filename: "tutorial02.xlsx")!
        let sheet = document.sheet()!
        var row = 0
        let col = 0
        var i = 0

        let bold = document.format()!
        bold.set(.bold)

        let money = document.format()!
        money.set(.number("$#,##0"))

        sheet.write(.string("Item"), location: Location(row, col), format: bold)
        sheet.write(.string("Cost"), location: Location(row, col + 1), format: bold)

        for expense in expenses {
            row = i + 1
            sheet.write(.string(expense.name), location: Location(row, col))
            sheet.write(.number(expense.cost), location: Location(row, col + 1), format: money)
            i += 1
        }

        sheet.write(.string("Total"), location: Location(row + 1, col), format: bold)
        sheet.write(.formula("=SUM(B2:B5)"), location: Location(row + 1, col + 1), format: money)

        document.close()
    }

    func testTutorial03() throws {
        struct Expense {
            let name: String
            let cost: Double
            let date: Date
        }

        let expenses: [Expense] = [
            Expense(name: "Rent", cost: 1000, date: get_date(year: 2013, month: 1, day: 13)),
            Expense(name: "Gas", cost: 100, date: get_date(year: 2013, month: 1, day: 14)),
            Expense(name: "Food", cost: 300, date: get_date(year: 2013, month: 1, day: 16)),
            Expense(name: "Gym", cost: 50, date: get_date(year: 2013, month: 1, day: 20))
        ]

        let document = Document(filename: "tutorial03.xlsx")!
        let sheet = document.sheet()!
        var row = 0
        let col = 0
        var i = 0

        let bold = document.format()!
        bold.set(.bold)

        let money = document.format()!
        money.set(.number("$#,##0"))

        let dateFormat = document.format()!
        dateFormat.set(.number("mmmm d yyyy"))

        sheet.column(width: 15, first: 0, last: 0)

        sheet.write(.string("Item"), location: Location(row, col), format: bold)
        sheet.write(.string("Cost"), location: Location(row, col + 1), format: bold)

        for expense in expenses {
            row = i + 1
            sheet.write(.string(expense.name), location: Location(row, col))
            sheet.write(.date(expense.date), location: Location(row, col + 1), format: dateFormat)
            sheet.write(.number(expense.cost), location: Location(row, col + 2), format: money)
            i += 1
        }

        sheet.write(.string("Total"), location: Location(row + 1, col), format: bold)
        sheet.write(.formula("=SUM(C2:C5)"), location: Location(row + 1, col + 2), format: money)

        document.close()
    }

    func testDateAndTimes01() throws {
        let number = 41333.5

        let document = Document(filename: "date_and_times01.xlsx")!
        let sheet = document.sheet()!

        let format = document.format()!
        format.set(.number("mmm d yyyy hh:mm AM/PM"))

        sheet.column(width: 20, first: 0, last: 0)

        sheet.write(.number(number), location: Location(0, 0))

        sheet.write(.number(number), location: Location(1, 0), format: format)

        document.close()
    }

    func testDateAndTimes02() throws {
        let date = get_date(year: 2013, month: 2, day: 28, hour: 12, minute: 0, second: 0)

        let document = Document(filename: "date_and_times02.xlsx")!
        let sheet = document.sheet()!

        let format = document.format()!
        format.set(.number("mmm d yyyy hh:mm AM/PM"))

        sheet.column(width: 20, first: 0, last: 0)

        sheet.write(.date(date), location: Location(0, 0))
        sheet.write(.date(date), location: Location(1, 0), format: format)

        document.close()
    }

    func testDateAndTimes03() throws {
        let document = Document(filename: "date_and_times03.xlsx")!
        let sheet = document.sheet()!

        let format = document.format()!
        format.set(.number("mmm d yyyy hh:mm AM/PM"))

        sheet.column(width: 20, first: 0, last: 0)

        sheet.write(.unix(0), location: Location(0, 0), format: format)
        sheet.write(.unix(1577836800), location: Location(1, 0), format: format)
        sheet.write(.unix(-2208988800), location: Location(2, 0), format: format)

        document.close()
    }

    func testHyperlinks() throws {
        let document = Document(filename: "hyperlinks.xlsx")!
        let sheet = document.sheet()!

        let defaultURLFormat = document.defaultURLFormat!

        let errorFormat = document.format()!
        errorFormat.set(.underline(.single))
        errorFormat.set(.fontColor(.red))

        sheet.column(width: 30, first: 0, last: 0)

        sheet.write(.url("http://libxlsxwriter.github.io"), location: Location(0, 0))

        sheet.write(.url("http://libxlsxwriter.github.io"), location: Location(2, 0))
        sheet.write(.string("Read the documentation."), location: Location(2, 0), format: defaultURLFormat)

        sheet.write(.url("http://libxlsxwriter.github.io"), location: Location(4, 0), format: errorFormat)

        sheet.write(.url("mailto:jmcnamara@cpan.org"), location: Location(6, 0))

        sheet.write(.url("mailto:jmcnamara@cpan.org"), location: Location(8, 0))
        sheet.write(.string("Drop me a line."), location: Location(8, 0), format: defaultURLFormat)

        document.close()
    }

    func testRichStrings() throws {
        let document = Document(filename: "rich_strings.xlsx")!
        let sheet = document.sheet()!

        let bold = document.format()!
        bold.set(.bold)

        let italic = document.format()!
        italic.set(.italic)

        let red = document.format()!
        red.set(.fontColor(.red))

        let blue = document.format()!
        blue.set(.fontColor(.blue))

        let center = document.format()!
        center.set(.alignment(.center))

        let superscript = document.format()!
        superscript.set(.fontScript(.superscript))

        sheet.column(width: 30, first: 0, last: 0)

        let fragment11 = RichString("This is ")
        let fragment12 = RichString("bold", format: bold)
        let fragment13 = RichString(" and this is ")
        let fragment14 = RichString("italic", format: italic)

        sheet.write(.richString([fragment11, fragment12, fragment13, fragment14]), location: Location("A1"))

        let fragment21 = RichString("This is ")
        let fragment22 = RichString("red", format: red)
        let fragment23 = RichString(" and this is ")
        let fragment24 = RichString("blue", format: blue)

        sheet.write(.richString([fragment21, fragment22, fragment23, fragment24]), location: Location("A3"))

        let fragment31 = RichString("Some ")
        let fragment32 = RichString("bold text", format: bold)
        let fragment33 = RichString(" centered")

        sheet.write(.richString([fragment31, fragment32, fragment33]), location: Location("A5"), format: center)

        let fragment41 = RichString("j =k", format: italic)
        let fragment42 = RichString("(n-1)", format: superscript)

        sheet.write(.richString([fragment41, fragment42]), location: Location("A7"), format: center)

        document.close()
    }

    func testArrayFormula() throws {
        let document = Document(filename: "array_formula.xlsx")!
        let sheet = document.sheet()!

        sheet.write(.number(500), location: Location(0, 1))
        sheet.write(.number(10), location: Location(1, 1))
        sheet.write(.number(1), location: Location(4, 1))
        sheet.write(.number(2), location: Location(5, 1))
        sheet.write(.number(3), location: Location(6, 1))

        sheet.write(.number(300), location: Location(0, 2))
        sheet.write(.number(15), location: Location(1, 2))
        sheet.write(.number(20234), location: Location(4, 2))
        sheet.write(.number(21003), location: Location(5, 2))
        sheet.write(.number(10000), location: Location(6, 2))

        sheet.write(.formula("{=SUM(B1:C1*B2:C2)}"), range: Range(0, 0, 0, 0))
        sheet.write(.formula("{=SUM(B1:C1*B2:C2)}"), range: Range("A2:A2"))
        sheet.write(.formula("{=TREND(C5:C7,B5:B7)}"), range: Range(4, 0, 6, 0))

        document.close()
    }

    func testUTF8() throws {
        let document = Document(filename: "utf8.xlsx")!
        let sheet = document.sheet()!

        sheet.write(.string("Это фраза на русском!"), location: Location(2, 1))

        document.close()
    }

    func testMergeRange() throws {
        let document = Document(filename: "merge_range.xlsx")!
        let sheet = document.sheet()!

        let mergeFormat = document.format()!

        mergeFormat.set(.alignment(.center))
        mergeFormat.set(.alignment(.verticalCenter))
        mergeFormat.set(.bold)
        mergeFormat.set(.backgroundColor(.yellow))
        mergeFormat.set(.border(.thin))

        sheet.column(width: 12, first: 1, last: 3)
        sheet.row(height: 30, row: 3)
        sheet.row(height: 30, row: 6)
        sheet.row(height: 30, row: 7)

        sheet.merge("Merged Range", range: Range(3, 1, 3, 3), format: mergeFormat)
        sheet.merge("Merged Range", range: Range(6, 1, 7, 3), format: mergeFormat)

        document.close()
    }

    func testMergeRichString() throws {
        let document = Document(filename: "merge_rich_string.xlsx")!
        let sheet = document.sheet()!

        let mergeFormat = document.format()!
        mergeFormat.set(.alignment(.center))
        mergeFormat.set(.alignment(.verticalCenter))
        mergeFormat.set(.border(.thin))

        let red = document.format()!
        red.set(.fontColor(.red))

        let blue = document.format()!
        blue.set(.fontColor(.blue))

        let fragment1 = RichString("This is ")
        let fragment2 = RichString("red", format: red)
        let fragment3 = RichString(" and this is ")
        let fragment4 = RichString("blue", format: blue)

        sheet.merge("", range: Range(1, 1, 4, 3), format: mergeFormat)

        sheet.write(.richString([fragment1, fragment2, fragment3, fragment4]), location: Location(1, 1), format: mergeFormat)

        document.close()
    }

    func testConditionalFormatSimple() throws {
        let document = Document(filename: "conditional_format_simple.xlsx")!
        let sheet = document.sheet()!

        sheet.write(.number(34), location: Location("B1"))
        sheet.write(.number(32), location: Location("B2"))
        sheet.write(.number(31), location: Location("B3"))
        sheet.write(.number(35), location: Location("B4"))
        sheet.write(.number(36), location: Location("B5"))
        sheet.write(.number(30), location: Location("B6"))
        sheet.write(.number(38), location: Location("B7"))
        sheet.write(.number(38), location: Location("B8"))
        sheet.write(.number(32), location: Location("B9"))

        let customFormat = document.format()!
        customFormat.set(.fontColor(.red))

        let conditionalFormat = ConditionalFormat(
            type: .cell,
            criteria: .lessThan,
            value: 33,
            format: customFormat
        )

        sheet.set(conditionalFormat, range: Range("B1:B9"))

        document.close()
    }

}

func get_date(year: Int, month: Int, day: Int, hour: Int = 0, minute: Int = 0, second: Int = 0) -> Date {
    Calendar.current.date(from: DateComponents(year: year, month: month, day: day, hour: hour, minute: minute, second: second))!
}
