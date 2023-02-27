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
  module Loader
    module Jrxml

      class TextField < StaticText

        attr_accessor :text_field_expression
        attr_accessor :is_stretch_with_overflow
        attr_accessor :is_blank_when_null
        attr_accessor :evaluation_time
        attr_accessor :pattern
        attr_accessor :pattern_expression
        attr_reader   :report_element

        def initialize (binding:)
          super(text: nil)
          @text_field_expression     = nil
          @is_blank_when_null        = nil
          @is_stretch_with_overflow  = false
          @evaluation_time           = nil
          if nil != binding
            @pattern                   = binding[:pattern]
            @pattern_expression        = binding[:pattern_expression]
            @report_element.properties = binding[:properties]
            @report_element.print_when_expression = binding[:printWhenExpression]
          else
            @pattern                   = nil
            @pattern_expression        = nil
            @report_element.properties = nil
          end
        end

        def attributes
          rv = Hash.new
          rv['isStretchWithOverflow'] = true if @is_stretch_with_overflow
          rv['pattern']               = @pattern unless @pattern.nil?
          rv['isBlankWhenNull']       = @is_blank_when_null unless @is_blank_when_null.nil?
          rv['evaluationTime']        = @evaluation_time unless @evaluation_time.nil?
          return rv
        end

        def to_xml (a_node)
          Nokogiri::XML::Builder.with(a_node) do |xml|
            xml.textField(attributes)
          end
          @report_element.to_xml(a_node.children.last)
          @box.to_xml(a_node.children.last) unless @box.nil?
          if nil != @text_field_expression && @text_field_expression.length > 0
            Nokogiri::XML::Builder.with(a_node.children.last) do |xml|
              xml.textFieldExpression {
                xml.cdata(@text_field_expression)
              }
            end
          end
          if nil != @pattern_expression && @pattern_expression.length > 0
            Nokogiri::XML::Builder.with(a_node.children.last) do |xml|
              xml.patternExpression {
                xml.cdata(@pattern_expression)
              }
            end
          end
        end

      end

    end
  end
end
