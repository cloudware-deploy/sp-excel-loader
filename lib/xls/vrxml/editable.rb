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

    attr_reader :styles

    class Editable

      def initialize()
          @styles = Hash.new
      end

      def load()

        default = Style.new("EditableTextField")
        default.mode      ="Opaque"
        default.forecolor ="#000000"
        default.backcolor ="#D2EAF0"
        add_style(style: default)

        invalid = Style.new("EditableTextFieldInvalidContent")
        invalid.style ="EditableTextField"
        invalid.box   = bottom_box("#E44A2C")
        add_style(style: invalid)

        focused = Style.new("EditableFocusedTextField")
        focused.mode    = "Opaque"
        focused.forecolor = "#808080"
        focused.backcolor = "#F7F2E1"
        add_style(style: focused)

        focused_invalid = Style.new("EditableFocusedInvalidContentTextField")
        focused_invalid.style = "EditableFocusedTextField"
        focused_invalid.box   = bottom_box("#E44A2C")
        add_style(style: focused_invalid)

        disabled = Style.new("EditableDisabledTextField")
        disabled.mode      = "Opaque"
        disabled.forecolor = "#C7C7C7"
        disabled.backcolor = "#F2F2F2"
        disabled.box = bottom_box("#000000", 1, "Dashed")
        add_style(style: disabled)
        
      end

      def set_style(name:, style:)
        @styles[name] = style.clone
        @styles[name].name = name
      end

      def styles_to_xml(node:)
        Nokogiri::XML::Builder.with(node) do |xml|
          xml.comment(" EDITABLE STYLES ")
          @styles.each do | _, style |
            style.to_xml(node)
          end
        end
      end

      private

      def add_style(style:)
        @styles[style.name] = style
      end

      def bottom_box(a_line_color, a_line_width=1, a_line_style="Solid")
        box                       = Box.new
        box.bottom_pen            = BottomPen.new
        box.bottom_pen.line_width = a_line_width
        box.bottom_pen.line_style = a_line_style
        box.bottom_pen.line_color = a_line_color
        box
      end

    end # Editable

  end # of module 'Vrxml'
end # of module 'Xls'

