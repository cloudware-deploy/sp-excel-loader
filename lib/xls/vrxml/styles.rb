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

    class Styles < Stylable

      def to_xml(node:)
        Nokogiri::XML::Builder.with(node) do |xml|
          xml.comment(" #{@name.upcase} ")
          @styles.each do | name, style |
            style.to_xml(node)
          end
        end
      end

      def self.parse(workbook:)
        styles = {}
        workbook.worksheets.each do |ws|
          if 'Styles' != ws.sheet_name
            next
          end
          style = Styles.new(name: ws.sheet_name)
          styles[style.name] = style
          for row in ws.dimension.ref.row_range
            if nil == ws[row] || nil == ws[row][0] || nil == ws[row][0].value
              next
            end
            tag = ws[row][0].value.to_s.strip
            case tag
            when /Style:.+/i
              if nil == style
                raise "Missing style object!"
              end
              style.add(style: xf_to_style(named: tag.split(':')[1].strip, workbook: workbook, style_index: ws[row][2].style_index))
            end
          end
          break
        end
        return styles
      end

    end # of class 'Theme'

  end # of module 'Vrxml'
end # of module 'Xls'

  