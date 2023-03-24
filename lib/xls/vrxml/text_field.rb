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

    class TextField < StaticText

      attr_accessor :text_field_expression
      attr_accessor :is_stretch_with_overflow
      attr_accessor :is_blank_when_null
      attr_accessor :evaluation_time
      attr_accessor :pattern
      attr_accessor :pattern_expression
      attr_reader   :report_element

      def initialize (binding:, cell: nil, text_field_expression: nil, pattern: nil, tracking: nil)
        super(text: nil)
        @text_field_expression     = nil
        @is_blank_when_null        = nil
        @is_stretch_with_overflow  = false
        @evaluation_time           = nil
        if nil != binding
          @pattern                     = pattern || binding[:pattern]
          @pattern_expression          = binding[:patternExpression]
          @is_stretch_with_overflow    = binding[:is_stretch_with_overflow] || binding[:isStretchWithOverflow] || ( binding.include?(:textAdjust) ? 'StretchHeight' == binding[:textAdjust] :false )
          @report_element.properties   = binding[:properties]
          @report_element.stretch_type = binding[:stretch_type] || binding[:stretchType]
        else
          @pattern                   = nil
          @pattern_expression        = nil
          @report_element.properties = nil
        end
        @text_field_expression = text_field_expression || binding[:text_field_expression] || binding[:textFieldExpression]
        @cell                  = cell
        @tracking              = tracking
      end

      def attributes
        rv = Hash.new
        rv[:isStretchWithOverflow] = true if @is_stretch_with_overflow
        rv[:pattern]               = @pattern unless @pattern.nil?
        rv[:isBlankWhenNull]       = @is_blank_when_null unless @is_blank_when_null.nil?
        rv[:evaluationTime]        = @evaluation_time unless @evaluation_time.nil?
        return rv
      end

      def to_xml (a_node)
        Nokogiri::XML::Builder.with(a_node) do |xml|
          if nil != @cell
            xml.comment(" #{@cell[:name] || @cell[:ref] || ''}#{@tracking ? " #{@tracking}" : '' } ")
          end  
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

  end # of module 'Vrxml'
end # of module 'Xls'
