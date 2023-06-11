//
//  ConditionalFormat.swift
//  
//
//  Created by Lucas Gladding on 2023-06-11.
//

import Foundation
import libxlsxwriter

public struct ConditionalFormat {
    public enum ConditionalFormatType {
        case none
        case cell

        var lxwConditionalFormatType: UInt8? {
            let options: [ConditionalFormatType: lxw_conditional_format_types] = [
                .none: LXW_CONDITIONAL_TYPE_NONE,
                .cell: LXW_CONDITIONAL_TYPE_CELL
            ]
            guard let option = options[self] else {
                return nil
            }
            return UInt8(option.rawValue)
        }
    }

    public enum ConditionalFormatCriteria {
        case lessThan

        var lxwConditionalFormatCriteria: UInt8? {
            let options: [ConditionalFormatCriteria: lxw_conditional_criteria] = [
                .lessThan: LXW_CONDITIONAL_CRITERIA_LESS_THAN
            ]
            guard let option = options[self] else {
                return nil
            }
            return UInt8(option.rawValue)
        }
    }

    let type: ConditionalFormatType
    let criteria: ConditionalFormatCriteria
    let value: Double
    let format: Format

    public init(type: ConditionalFormatType, criteria: ConditionalFormatCriteria, value: Double, format: Format) {
        self.type = type
        self.criteria = criteria
        self.value = value
        self.format = format
    }

    var lxwConditionalFormat: lxw_conditional_format {
        var format = lxw_conditional_format()
        format.type = self.type.lxwConditionalFormatType!
        format.criteria = self.criteria.lxwConditionalFormatCriteria!
        format.value = self.value
        format.format = self.format.lxw_format

        return format
    }
}
