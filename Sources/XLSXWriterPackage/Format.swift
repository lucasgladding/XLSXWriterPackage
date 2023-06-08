//
//  Format.swift
//  XLSXWriterPackage
//
//  Created by Lucas Gladding on 2022-12-03.
//

import Foundation
import libxlsxwriter

public class Format {
    public enum Alignment {
        case center
    }

    public enum Color {
        case black
        case blue
        case cyan
        case gray
        case green
        case lime
        case magenta
        case navy
        case orange
        case pink
        case purple
        case red
        case silver
        case white
        case yellow
    }

    public enum Script {
        case superscript
    }

    public enum Underline {
        case double
        case single
    }

    public enum Options {
        case alignment(Alignment)
        case bold
        case fontColor(Color)
        case fontScript(Script)
        case italic
        case number(String)
        case underline(Underline)
    }

    let lxw_format: UnsafeMutablePointer<lxw_format>

    init(lxw_format: UnsafeMutablePointer<lxw_format>) {
        self.lxw_format = lxw_format
    }

    public func set(_ options: Options) {
        switch options {
        case .alignment(let input):
            if let alignment = get_lxw_alignment(from: input) {
                format_set_align(self.lxw_format, alignment)
            }
        case .bold:
            format_set_bold(self.lxw_format)
        case .fontColor(let input):
            if let fontColor = get_lxw_color(from: input) {
                format_set_font_color(self.lxw_format, fontColor)
            }
        case .fontScript(let input):
            if let fontScript = get_lxw_script(from: input) {
                format_set_font_script(self.lxw_format, fontScript)
            }
        case .italic:
            format_set_italic(self.lxw_format)
        case .number(let input):
            input.withCString { format_set_num_format(self.lxw_format, $0) }
        case .underline(let input):
            if let underline = get_lxw_underline(from: input) {
                format_set_underline(self.lxw_format, underline)
            }
        }
    }

    private func get_lxw_alignment(from alignment: Alignment) -> UInt8? {
        let styles: [Alignment: lxw_format_alignments] = [
            .center: LXW_ALIGN_CENTER
        ]
        guard let style = styles[alignment] else {
            return nil
        }
        return UInt8(style.rawValue)
    }

    private func get_lxw_color(from color: Color) -> UInt32? {
        let styles: [Color: lxw_defined_colors] = [
            .black: LXW_COLOR_BLACK,
            .blue: LXW_COLOR_BLUE,
            .cyan: LXW_COLOR_CYAN,
            .gray: LXW_COLOR_GRAY,
            .green: LXW_COLOR_GREEN,
            .lime: LXW_COLOR_LIME,
            .magenta: LXW_COLOR_MAGENTA,
            .navy: LXW_COLOR_NAVY,
            .orange: LXW_COLOR_ORANGE,
            .pink: LXW_COLOR_PINK,
            .purple: LXW_COLOR_PURPLE,
            .red: LXW_COLOR_RED,
            .silver: LXW_COLOR_SILVER,
            .white: LXW_COLOR_WHITE,
            .yellow: LXW_COLOR_YELLOW
        ]
        guard let style = styles[color] else {
            return nil
        }
        return UInt32(style.rawValue)
    }

    private func get_lxw_script(from script: Script) -> UInt8? {
        let styles: [Script: lxw_format_scripts] = [
            .superscript: LXW_FONT_SUPERSCRIPT
        ]
        guard let style = styles[script] else {
            return nil
        }
        return UInt8(style.rawValue)
    }

    private func get_lxw_underline(from underline: Underline) -> UInt8? {
        let styles: [Underline: lxw_format_underlines] = [
            .double: LXW_UNDERLINE_DOUBLE,
            .single: LXW_UNDERLINE_SINGLE
        ]
        guard let style = styles[underline] else {
            return nil
        }
        return UInt8(style.rawValue)
    }
}
