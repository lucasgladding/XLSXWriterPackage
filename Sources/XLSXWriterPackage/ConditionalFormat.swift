//
//  ConditionalFormat.swift
//  
//
//  Created by Lucas Gladding on 2023-06-11.
//

import Foundation
import libxlsxwriter

public struct ConditionalFormat {
    public enum ConditionalType {
        case none
        case cell
    }

    public enum Criteria {
        case lessThan
    }

    let type: ConditionalType
    let criteria: Criteria
    let value: Double
    let format: Format

    public init(type: ConditionalType, criteria: Criteria, value: Double, format: Format) {
        self.type = type
        self.criteria = criteria
        self.value = value
        self.format = format
    }
}

extension ConditionalFormat {
    var lxwConditionalFormat: lxw_conditional_format {
        var format = lxw_conditional_format()
        format.type = lxwConditionalFormatType!
        format.criteria = lxwConditionalFormatCriteria!
        format.value = self.value
        format.format = self.format.lxw_format
        return format
    }

    var lxwConditionalFormatType: UInt8? {
        let options: [ConditionalType: lxw_conditional_format_types] = [
            .none: LXW_CONDITIONAL_TYPE_NONE,
            .cell: LXW_CONDITIONAL_TYPE_CELL
        ]
        guard let option = options[self.type] else {
            return nil
        }
        return UInt8(option.rawValue)
    }

    var lxwConditionalFormatCriteria: UInt8? {
        let options: [Criteria: lxw_conditional_criteria] = [
            .lessThan: LXW_CONDITIONAL_CRITERIA_LESS_THAN
        ]
        guard let option = options[self.criteria] else {
            return nil
        }
        return UInt8(option.rawValue)
    }
}