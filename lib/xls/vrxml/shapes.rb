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

    ##################
    # GraphicElement #
    ##################
    class GraphicElement

      attr_accessor :pen

      def initialize()
        @pen = Pen.new
      end

      def attributes
        rv = Hash.new
        return rv
      end

      def to_xml(a_node)
        Nokogiri::XML::Builder.with(a_node) do |xml|
          xml.graphicElement(attributes)
        end
        @pen.to_xml(a_node.children.last)
      end

    end # of class 'GraphicElement'

    ##################
    # Shape          #
    ##################
    class Shape

      def initialize(cell:, tracking:, comment: nil)
        @report_element  = ReportElement.new
        @graphic_element = GraphicElement.new
        @cell            = cell
        @tracking        = tracking
        @comment         = comment
      end

      def x=(value)
        @report_element.x = value
      end

      def y=(value)
        @report_element.y = value
      end

      def width=(value)
        @report_element.width = value
      end

      def height=(value)
        @report_element.height = value
      end

      def forecolor=(value)
        @report_element.forecolor = value
      end

      def backcolor=(value)
        @report_element.backcolor = value
      end

      def attributes
        rv = Hash.new
        return rv
      end
      
      def to_xml(a_node)
        raise "Not Implemented!"
      end

    end # of class 'Shape'

    ##################
    # Rectangle      #
    ##################

    class Rectangle < Shape

      attr_accessor :radius

      def initialize(cell:, tracking:, comment: nil)
        super(cell: cell, tracking: tracking, comment: comment)
        @radius = nil
      end

      def attributes
        rv = super
        rv['radius'] = @radius if nil != @radius
        return rv
      end

      def to_xml(a_node)
        Nokogiri::XML::Builder.with(a_node) do |xml|
          if nil != @cell
            xml.comment(" #{@comment || ''} #{@cell[:name] || @cell[:ref] || ''}#{@tracking ? " #{@tracking}" : '' } ")
          end
          xml.rectangle(attributes)
        end
        @report_element.to_xml(a_node.children.last)
        @graphic_element.to_xml(a_node.children.last)
      end

    end # of class 'Rectangle'

  end # of module 'Vrxml'
end # of module 'Xls'