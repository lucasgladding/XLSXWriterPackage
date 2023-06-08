//
//  Book.swift
//  XLSXWriterPackage
//
//  Created by Lucas Gladding on 2022-12-03.
//

import Foundation
import libxlsxwriter

public class Document {
    public let path: String

    private let lxw_workbook: UnsafeMutablePointer<lxw_workbook>

    init(path: String) {
        self.path = path
        self.lxw_workbook = path.withCString { workbook_new($0) }
    }

    public func sheet(name: String? = nil) -> Sheet? {
        guard let lxw_worksheet = get_lxw_worksheet(name: name) else {
            return nil
        }

        return Sheet(lxw_worksheet: lxw_worksheet)
    }

    private func get_lxw_worksheet(name: String? = nil) -> UnsafeMutablePointer<lxw_worksheet>? {
        if let name = name {
            return name.withCString { workbook_add_worksheet(self.lxw_workbook, $0) }
        } else {
            return workbook_add_worksheet(self.lxw_workbook, nil)
        }
    }

    public func format() -> Format? {
        guard let lxw_format = workbook_add_format(self.lxw_workbook) else {
            return nil
        }

        return Format(lxw_format: lxw_format)
    }

    public var defaultURLFormat: Format? {
        guard let lxw_format = workbook_get_default_url_format(self.lxw_workbook) else {
            return nil
        }

        return Format(lxw_format: lxw_format)
    }

    public func close() {
        workbook_close(self.lxw_workbook)
    }
}

extension Document {
    convenience public init(filename: String) {
        let documentURL = try! FileManager.default.documentURL(filename: filename)
        self.init(path: documentURL.path)
    }
}
