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

    public func set(width: Double, first: Int, last: Int) {
        worksheet_set_column(self.lxw_worksheet, UInt16(first), UInt16(last), width, nil)
    }

    public func write(_ content: Content, location: Location, format: Format? = nil) {
        var error: lxw_error?

        let r = UInt32(location.row)
        let c = UInt16(location.col)
        let format = get_lxw_format(from: format)

        switch content {
        case .date(let date):
            var datetime = get_lxw_datetime(from: date)
            worksheet_write_datetime(self.lxw_worksheet, r, c, &datetime, format)
        case .formula(let formula):
            error = formula.withCString { worksheet_write_formula(self.lxw_worksheet, r, c, $0, format) }
        case .image(let image):
            error = image.withCString { worksheet_insert_image(self.lxw_worksheet, r, c, $0) }
        case .number(let number):
            error = worksheet_write_number(self.lxw_worksheet, r, c, number, format)
        case .richString(let richString):
            error = write(richString: richString, row: r, col: c, format: format)
        case .string(let string):
            error = string.withCString { worksheet_write_string(self.lxw_worksheet, r, c, $0, format) }
        case .unix(let unix):
            error = worksheet_write_unixtime(self.lxw_worksheet, r, c, Int64(unix), format)
        case .url(let url):
            error = url.withCString { worksheet_write_url(self.lxw_worksheet, r, c, $0, format) }
        }

        if let errorString = errorString(from: error) {
            print("Error writing location \(errorString)")
        }
    }

    public func write(_ content: Content, range: Range, format: Format? = nil) {
        var error: lxw_error?

        let r1 = UInt32(range.location1.row)
        let c1 = UInt16(range.location1.col)
        let r2 = UInt32(range.location2.row)
        let c2 = UInt16(range.location2.col)
        let format = get_lxw_format(from: format)

        switch content {
        case .formula(let formula):
            error = formula.withCString { cString in
                worksheet_write_array_formula(self.lxw_worksheet, r1, c1, r2, c2, cString, format)
            }
        default:
            print("Unsupported content")
        }

        if let errorString = errorString(from: error) {
            print("Error writing range \(errorString)")
        }
    }

    private func write(
        richString: [RichString],
        row: UInt32,
        col: UInt16,
        format: UnsafeMutablePointer<lxw_format>? = nil
    ) -> lxw_error? {
        var error: lxw_error?

        var richStringTupleArray = richString.map { richString in
            get_lxw_rich_string_tuple(from: richString)
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

    private func get_lxw_format(from format: Format?) -> UnsafeMutablePointer<lxw_format>? {
        format?.lxw_format
    }

    private func get_lxw_datetime(from date: Date) -> lxw_datetime {
        let components = Calendar.current.dateComponents([.year, .month, .day, .hour, .minute, .second], from: date)
        return lxw_datetime(
            year: Int32(components.year!),
            month: Int32(components.month!),
            day: Int32(components.day!),
            hour: Int32(components.hour!),
            min: Int32(components.minute!),
            sec: Double(components.second!)
        )
    }

    private func get_lxw_rich_string_tuple(from richString: RichString) -> lxw_rich_string_tuple {
        lxw_rich_string_tuple(
            format: get_lxw_format(from: richString.format),
            string: richString.string.cString
        )
    }

    private func errorString(from error: lxw_error?) -> String? {
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
