//
//  Sheet.swift
//  XLSXWriterPackage
//
//  Created by Lucas Gladding on 2022-12-03.
//

import Foundation
import libxlsxwriter

public class Sheet {
    public enum Content {
        case date(Date)
        case formula(String)
        case image(String)
        case number(Double)
        case richString([RichString])
        case string(String)
        case unix(Int)
        case url(String)
    }

    private let lxw_worksheet: UnsafeMutablePointer<lxw_worksheet>

    init(lxw_worksheet: UnsafeMutablePointer<lxw_worksheet>) {
        self.lxw_worksheet = lxw_worksheet
    }

    public func format(_ format: ConditionalFormat, range: Range) {
        let r1 = UInt32(range.location1.row)
        let c1 = UInt16(range.location1.col)
        let r2 = UInt32(range.location2.row)
        let c2 = UInt16(range.location2.col)
        var lxwConditionalFormat = format.lxwConditionalFormat

        worksheet_conditional_format_range(self.lxw_worksheet, r1, c1, r2, c2, &lxwConditionalFormat)
    }

    public func height(_ height: Double, row: Int, format: Format? = nil) {
        worksheet_set_row(self.lxw_worksheet, UInt32(row), Double(height), format?.lxw_format)
    }

    public func merge(_ string: String, range: Range, format: Format? = nil) {
        let r1 = UInt32(range.location1.row)
        let c1 = UInt16(range.location1.col)
        let r2 = UInt32(range.location2.row)
        let c2 = UInt16(range.location2.col)
        let lxwFormat = format?.lxw_format

        _ = string.withCString { worksheet_merge_range(self.lxw_worksheet, r1, c1, r2, c2, $0, lxwFormat) }
    }

    public func width(_ width: Double, first: Int, last: Int, format: Format? = nil) {
        let lxwFormat = format?.lxw_format
        worksheet_set_column(self.lxw_worksheet, UInt16(first), UInt16(last), width, lxwFormat)
    }

    public func write(_ content: Content, location: Location, format: Format? = nil) {
        var error: lxw_error?

        let r = UInt32(location.row)
        let c = UInt16(location.col)
        let lxwFormat = format?.lxw_format

        switch content {
        case .date(let input):
            var datetime = input.lxwDateTime
            worksheet_write_datetime(self.lxw_worksheet, r, c, &datetime, lxwFormat)
        case .formula(let input):
            error = input.withCString { worksheet_write_formula(self.lxw_worksheet, r, c, $0, lxwFormat) }
        case .image(let input):
            error = input.withCString { worksheet_insert_image(self.lxw_worksheet, r, c, $0) }
        case .number(let input):
            error = worksheet_write_number(self.lxw_worksheet, r, c, input, lxwFormat)
        case .richString(let input):
            error = writeRichString(input, row: r, col: c, format: lxwFormat)
        case .string(let input):
            error = input.withCString { worksheet_write_string(self.lxw_worksheet, r, c, $0, lxwFormat) }
        case .unix(let input):
            error = worksheet_write_unixtime(self.lxw_worksheet, r, c, Int64(input), lxwFormat)
        case .url(let input):
            error = input.withCString { worksheet_write_url(self.lxw_worksheet, r, c, $0, lxwFormat) }
        }

        if let errorString = getErrorString(from: error) {
            print("Error writing location \(errorString)")
        }
    }

    public func write(_ content: Content, range: Range, format: Format? = nil) {
        var error: lxw_error?

        let r1 = UInt32(range.location1.row)
        let c1 = UInt16(range.location1.col)
        let r2 = UInt32(range.location2.row)
        let c2 = UInt16(range.location2.col)
        let lxwFormat = format?.lxw_format

        switch content {
        case .formula(let input):
            error = input.withCString { cString in
                worksheet_write_array_formula(self.lxw_worksheet, r1, c1, r2, c2, cString, lxwFormat)
            }
        default:
            print("Unsupported content")
        }

        if let errorString = getErrorString(from: error) {
            print("Error writing range \(errorString)")
        }
    }

    private func writeRichString(
        _ content: [RichString],
        row: UInt32,
        col: UInt16,
        format: UnsafeMutablePointer<lxw_format>? = nil
    ) -> lxw_error? {
        var error: lxw_error?

        var richStringTupleArray = content.map { richString in
            richString.lxwRichStringTuple
        }

        richStringTupleArray.withUnsafeMutableBufferPointer { pointer1 in
            guard let baseAddress = pointer1.baseAddress else {
                return
            }

            var elements: [UnsafeMutablePointer<lxw_rich_string_tuple>?] = Array(baseAddress ..< baseAddress.advanced(by: pointer1.count))
            elements.append(nil)

            elements.withUnsafeMutableBufferPointer { pointer2 in
                error = worksheet_write_rich_string(self.lxw_worksheet, row, col, pointer2.baseAddress, format)
            }
        }

        return error
    }

    private func getErrorString(from error: lxw_error?) -> String? {
        guard let error = error else {
            return nil
        }
        guard error != LXW_NO_ERROR else {
            return nil
        }
        let cErrorString: UnsafeMutablePointer<CChar> = lxw_strerror(error)
        return String(cString: cErrorString)
    }
}
