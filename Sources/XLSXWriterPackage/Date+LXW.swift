//
//  Date+LXW.swift
//  
//
//  Created by Lucas Gladding on 2023-06-11.
//

import Foundation
import libxlsxwriter

extension Date {
    var lxwDateTime: lxw_datetime {
        let components = Calendar.current.dateComponents([.year, .month, .day, .hour, .minute, .second], from: self)
        return lxw_datetime(
            year: Int32(components.year!),
            month: Int32(components.month!),
            day: Int32(components.day!),
            hour: Int32(components.hour!),
            min: Int32(components.minute!),
            sec: Double(components.second!)
        )
    }
}
