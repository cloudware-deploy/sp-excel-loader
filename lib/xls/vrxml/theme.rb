# encoding: utf-8
#
# Copyright (c) 2011-2023 Cloudware S.A. All rights reserved.
#
# This file is part of xls2vrxml.
# Based on sp-excel-loader.
#
# xls2vrxml is free software: you can redistribute it and/or modify
# it under the terms of the GNU Affero General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# xls2vrxml is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU Affero General Public License
# along with xls2vrxml.  If not, see <http://www.gnu.org/licenses/>.
#

module Xls
  module Vrxml

    class Theme

      attr_reader :name
      attr_reader :styles

      def initialize(name:)
        @name   = name
        @styles = {}
      end

      def attributes
        rv = Hash.new
        rv['name']  = @name
        return rv
      end

      def to_xml(node:, theme: nil)
        Nokogiri::XML::Builder.with(node) do |xml|
          xml.comment(" Theme: #{@name} ")
          xml.theme(attributes) {
            @styles.each do | name, style |
              style.to_xml(node.children.last)
            end
          }
        end
      end

      def add(style:)
        @styles[style.name] = style
      end

      def self.parse(workbook:)
        themes = {}
        workbook.worksheets.each do |ws|
          if false == ws.sheet_name.start_with?('Theme')
            next
          end
          theme = nil
          for row in ws.dimension.ref.row_range
            if nil == ws[row] || nil == ws[row][0] || nil == ws[row][0].value
              next
            end
            tag = ws[row][0].value.to_s.strip
            case tag
            when /Theme:.+/i
              name = tag.split(':')[1].strip
              if nil != name && false == themes.include?(name)
                theme = Theme.new(name: ws.sheet_name)
                themes[name] = theme
              end
            when /Style:.+/i
              if nil == theme
                raise "Missing theme object!"
              end
              theme.add(style: xf_to_style(named: tag.split(':')[1].strip, workbook: workbook, style_index: ws[row][2].style_index))
            end
          end
        end
        return themes
      end

      def self.xf_to_style(named:, workbook:, style_index:)

        # create a style
        style = Style.new(named)

        # grab cell format
        xf = workbook.cell_xfs[style_index]

        # Format font
        if xf.apply_font == true
          xls_font = workbook.fonts[xf.font_id]

          # on PDF we only have one font
          style.font_name = 'DejaVu Sans Condensed'
          # if xls_font.name.val == 'Arial'
          #   style.font_name = 'DejaVu Sans Condensed'
          # else
          #   style.font_name = xls_font.name.val
          # end

          unless xls_font.color.nil?
            style.forecolor = convert_color(workbook: workbook, xls_color: xls_font.color)
          end

          style.font_size = xls_font.sz.val unless xls_font.sz.nil?
          style.is_bold   = true unless xls_font.b.nil?
          style.is_italic = true unless xls_font.i.nil?
        end

        # background
        if xf.apply_fill == true
          xls_fill = workbook.fills[xf.fill_id]
          if xls_fill.pattern_fill.pattern_type == 'solid'
            style.backcolor = convert_color(workbook: workbook, xls_color: xls_fill.pattern_fill.fg_color)
          end
        end

        # borders
        if xf.apply_border == true
          xls_border = workbook.borders[xf.border_id]

          if xls_border.outline != nil

            if xls_border.outline.style != nil
              style.box ||= Box.new
              style.box.left_pen  = LeftPen.new
              style.box.top_pen   = TopPen.new
              style.box.right_pen = RightPen.new
              style.box.bottom    = BottomPen.new
              apply_border_style(workbook: workbook, pen: style.box.left_pen  , xls_border_style: xls_border.outline)
              apply_border_style(workbook: workbook, pen: style.box.top_pen   , xls_border_style: xls_border.outline)
              apply_border_style(workbook: workbook, pen: style.box.right_pen , xls_border_style: xls_border.outline)
              apply_border_style(workbook: workbook, pen: style.box.bottom_pen, xls_border_style: xls_border.outline)
            end

          else

            if xls_border.left != nil && xls_border.left.style != nil
              style.box ||= Box.new
              style.box.left_pen = LeftPen.new
              apply_border_style(workbook: workbook, pen: style.box.left_pen, xls_border_style: xls_border.left)
            end

            if xls_border.top != nil && xls_border.top.style != nil
              style.box ||= Box.new
              style.box.top_pen = TopPen.new
              apply_border_style(workbook: workbook, pen: style.box.top_pen, xls_border_style: xls_border.top)
            end

            if xls_border.right != nil && xls_border.right.style != nil
              style.box ||= Box.new
              style.box.right_pen = RightPen.new
              apply_border_style(workbook: workbook, pen: style.box.right_pen, xls_border_style: xls_border.right)
            end

            if xls_border.bottom != nil && xls_border.bottom.style != nil
              style.box ||= Box.new
              style.box.bottom_pen = BottomPen.new
              apply_border_style(workbook: workbook, pen: style.box.bottom_pen, xls_border_style: xls_border.bottom)
            end

          end
        end

        # Alignment
        if xf.apply_alignment

          #byebug if style_index == 111
          unless xf.alignment.nil?
            case xf.alignment.horizontal
            when 'left', nil
              style.h_text_align ='Left'
            when 'center'
              style.h_text_align ='Center'
            when 'right'
              style.h_text_align ='Right'
            end

            case xf.alignment.vertical
            when 'top'
              style.v_text_align ='Top'
            when 'center'
              style.v_text_align ='Middle'
            when 'bottom', nil
              style.v_text_align ='Bottom'
            end

            # rotation
            case xf.alignment.text_rotation
            when nil
              style.rotation = nil
            when 0
              style.rotation = 'None'
            when 90
              style.rotation = 'Left'
            when 180
              style.rotation = 'UpsideDown'
            when 270
              style.rotation = 'Right'
            end
          end
        end

        return style

      end

      def self.apply_border_style(workbook:, pen:, xls_border_style:)
        case xls_border_style.style
        when 'thin'
          pen.line_width = 0.5
          pen.line_style = 'Solid'
        when 'medium'
          pen.line_width = 1.0
          pen.line_style = 'Solid'
        when 'dashed'
          pen.line_width = 1.0
          pen.line_style = 'Dotted'
        when 'dotted'
          pen.line_width = 0.5
          pen.line_style = 'Dotted'
        when 'thick'
          pen.line_width = 2.0
          pen.line_style = 'Solid'
        when 'double'
          pen.line_width = 0.5
          pen.line_style = 'Double'
        when 'hair'
          pen.line_width = 0.25
          pen.line_style = 'Solid'
        when 'mediumDashed'
          pen.line_width = 1.0
          pen.line_style = 'Dashed'
        when 'dashDot'
          pen.line_width = 0.5
          pen.line_style = 'Dashed'
        when 'mediumDashDot'
          pen.line_width = 1.0
          pen.line_style = 'Dashed'
        when 'dashDotDot'
          pen.line_width = 0.5
          pen.line_style = 'Dotted'
        when 'slantDashDot'
          pen.line_width = 0.5
          pen.line_style = 'Dotted'
        else
          pen.line_width = 1.0
          pen.line_style = 'Solid'
        end
        pen.line_color = convert_color(workbook: workbook, xls_color: xls_border_style.color)
      end

      def self.convert_color(workbook:, xls_color:)
        if xls_color.indexed.nil?
          if xls_color.theme != nil
            cs = workbook.theme.a_theme_elements.a_clr_scheme
            case xls_color.theme
            when 0
              return tint_theme_color(cs.a_lt1, xls_color.tint)
            when 1
              return tint_theme_color(cs.a_dk1, xls_color.tint)
            when 2
              return tint_theme_color(cs.a_lt2, xls_color.tint)
            when 3
              return tint_theme_color(cs.a_dk2, xls_color.tint)
            when 4
              return tint_theme_color(cs.a_accent1, xls_color.tint)
            when 5
              return tint_theme_color(cs.a_accent2, xls_color.tint)
            when 6
              return tint_theme_color(cs.a_accent3, xls_color.tint)
            when 7
              return tint_theme_color(cs.a_accent4, xls_color.tint)
            when 8
              return tint_theme_color(cs.a_accent5, xls_color.tint)
            when 9
              return tint_theme_color(cs.a_accent6, xls_color.tint)
            else
              return '#c0c0c0'
            end

          elsif xls_color.auto or xls_color.rgb.nil?
            return '#000000'
          else
            return '#' + xls_color.rgb[2..-1]
          end
        else
          return '#' + @@CT_IndexedColors[xls_color.indexed]
        end
      end

      def self.tint_theme_color(a_color, a_tint)
        color   = a_color.a_sys_clr.last_clr unless a_color.a_sys_clr.nil?
        color ||= a_color.a_srgb_clr.val
        r = color[0..1].to_i(16)
        g = color[2..3].to_i(16)
        b = color[4..5].to_i(16)
        unless a_tint.nil?
          if ( a_tint <  0 )
            a_tint = 1 + a_tint;
            r = r * a_tint
            g = g * a_tint
            b = b * a_tint
          else
            r = r + (a_tint * (255 - r))
            g = g + (a_tint * (255 - g))
            b = b + (a_tint * (255 - b))
          end
        end
        r = 255 if r > 255
        g = 255 if g > 255
        b = 255 if b > 255
        color = "#%02X%02X%02X" % [r, g, b]
        color
      end

    end # of class 'Theme'

  end # of module 'Vrxml'
end # of module 'Xls'

  