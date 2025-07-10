//
//  Format.swift
//  XLSXWriterPackage
//
//  Created by Lucas Gladding on 2022-12-03.
//

import Foundation
import libxlsxwriter

public class Format {
    public enum Options {
        case alignment(Alignment)
        case backgroundColor(Color)
        case bold
        case border(Border)
        case top(Border)
        case bottom(Border)
        case fontColor(Color)
        case fontScript(Script)
        case italic
        case number(String)
        case underline(Underline)
    }

    public enum Alignment {
        case left
        case center
        case verticalCenter

        var lxwAlignment: UInt8? {
            let options: [Alignment: lxw_format_alignments] = [
                .left: LXW_ALIGN_LEFT,
                .center: LXW_ALIGN_CENTER,
                .verticalCenter: LXW_ALIGN_VERTICAL_CENTER
            ]
            guard let option = options[self] else {
                return nil
            }
            return UInt8(option.rawValue)
        }
    }

    public enum Border {
        case thin
        case double

        var lxwBorder: UInt8? {
            let options: [Border: lxw_format_borders] = [
                .thin: LXW_BORDER_THIN,
                .double: LXW_BORDER_DOUBLE
            ]
            guard let option = options[self] else {
                return nil
            }
            return UInt8(option.rawValue)
        }
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

        var lxwColor: UInt32? {
            let options: [Color: lxw_defined_colors] = [
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
            guard let option = options[self] else {
                return nil
            }
            return UInt32(option.rawValue)
        }
    }

    public enum Script {
        case superscript

        var lxwScript: UInt8? {
            let options: [Script: lxw_format_scripts] = [
                .superscript: LXW_FONT_SUPERSCRIPT
            ]
            guard let option = options[self] else {
                return nil
            }
            return UInt8(option.rawValue)
        }
    }

    public enum Underline {
        case double
        case single

        var lxwUnderline: UInt8? {
            let options: [Underline: lxw_format_underlines] = [
                .double: LXW_UNDERLINE_DOUBLE,
                .single: LXW_UNDERLINE_SINGLE
            ]
            guard let option = options[self] else {
                return nil
            }
            return UInt8(option.rawValue)
        }
    }

    let lxw_format: UnsafeMutablePointer<lxw_format>

    init(lxw_format: UnsafeMutablePointer<lxw_format>) {
        self.lxw_format = lxw_format
    }

    public func set(_ options: Options) {
        switch options {
        case .alignment(let input):
            if let option = input.lxwAlignment {
                format_set_align(self.lxw_format, option)
            }
        case .backgroundColor(let input):
            if let option = input.lxwColor {
                format_set_bg_color(self.lxw_format, option)
            }
        case .border(let input):
            if let option = input.lxwBorder {
                format_set_border(self.lxw_format, option)
            }
        case .top(let input):
            if let option = input.lxwBorder {
                format_set_top(self.lxw_format, option)
            }
        case .bottom(let input):
            if let option = input.lxwBorder {
                format_set_bottom(self.lxw_format, option)
            }
        case .bold:
            format_set_bold(self.lxw_format)
        case .fontColor(let input):
            if let option = input.lxwColor {
                format_set_font_color(self.lxw_format, option)
            }
        case .fontScript(let input):
            if let option = input.lxwScript {
                format_set_font_script(self.lxw_format, option)
            }
        case .italic:
            format_set_italic(self.lxw_format)
        case .number(let input):
            input.withCString { format_set_num_format(self.lxw_format, $0) }
        case .underline(let input):
            if let option = input.lxwUnderline {
                format_set_underline(self.lxw_format, option)
            }
        }
    }
}
