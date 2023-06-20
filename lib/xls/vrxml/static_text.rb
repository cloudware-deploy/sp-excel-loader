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

    class StaticText

      attr_accessor :report_element
      attr_accessor :text
      attr_accessor :style
      attr_accessor :box
      attr_accessor :attributes

      # custom
      attr_accessor :theme_style

      def initialize(binding:, text:, cell: nil, tracking: nil)
        @binding        = Marshal.load(Marshal.dump(binding))
        @report_element = ReportElement.new
        @text           = text
        @box            = nil
        @attributes     = nil
        @cell           = cell
        @tracking       = tracking
        if nil != binding
          if binding[:theme_style]
            @report_element.theme_style = binding[:theme_style]
          end
          @report_element.properties   = binding[:properties]
          if binding[:stretch_type] || binding[:stretchType]
            @report_element.stretch_type = binding[:stretch_type] || binding[:stretchType]
          end
          @report_element.print_when_expression = binding[:printWhenExpression]
        end
      end

      def to_xml (a_node)
        # -
        background_to_xml(a_node)
        # -
        Nokogiri::XML::Builder.with(a_node) do |xml|
          if nil != @cell
            xml.comment(" #{@cell[:name] || @cell[:ref] || ''}#{@tracking ? " #{@tracking}" : '' } ")
          end
          xml.staticText(attributes)
        end
        @report_element.to_xml(a_node.children.last)
        box_to_xml(a_node.children.last)
        Nokogiri::XML::Builder.with(a_node.children.last) do |xml|
          xml.text_ {
            xml.cdata(@text)
          }
        end
        # -
        foreground_to_xml(a_node)
      end

      def box_to_xml(a_node)
        if nil != @box 
          if nil != @binding && nil != @binding[:padding]
            if @binding[:padding].is_a?(Hash)
              @binding[:padding].each do | k, v |
                _attr = k.to_s.to_underscore
                if @box.respond_to?(_attr.to_sym)
                  @box.send("#{_attr}=", v)
                end
              end
            else
              @box.padding = @binding[:padding].to_i
            end
          end
          @box.to_xml(a_node)
        end
      end

      #
      # Add a 'shape' node as 'background'.
      #
      # @param node Node where to append this 'shape'
      #.
      def background_to_xml(a_node)
        shape_to_xml(as: :background, node: a_node)        
      end

      #
      # Add a 'shape' node as 'foreground'.
      #
      # @param node Node where to append this 'shape'.
      #
      def foreground_to_xml(a_node)
        shape_to_xml(as: :foreground, node: a_node)
      end

      private

      #
      # Add a 'shape' node to a node.
      #
      # @param as   One of ':background', ':foreground'
      # @param node Node where to append this 'shape'.
      #
      def shape_to_xml(as:, node:)
        if nil != @binding && nil != @binding[as]
          if nil != @binding[as][:shape]
            case @binding[as][:shape]
            when 'rectangle'
              r = Rectangle.new(cell: @cell, tracking: @tracking, comment: "AS #{as.to_s.upcase} OF")
              r.x      = @report_element.x
              r.y      = @report_element.y
              r.width  = @binding[as][:width]  || @report_element.width
              r.height = @binding[as][:height] || @report_element.height
              r.radius = @binding[as][:radius]
              r.to_xml(node)
            else
            end
          end
        end
      end

    end # class 'StaticText'

  end # of module 'Vrxml'
end # of module 'Xls'
