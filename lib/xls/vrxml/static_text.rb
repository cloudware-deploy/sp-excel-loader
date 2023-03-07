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

      def initialize(text:, ref: nil, tracking: nil)
        @report_element = ReportElement.new
        @text           = text
        @box            = nil
        @attributes     = nil
        @cell_reference = ref
        @tracking       = tracking
      end

      def to_xml (a_node)
        Nokogiri::XML::Builder.with(a_node) do |xml|
          if nil != @cell_reference && @cell_reference.length > 0
            xml.comment(" #{@cell_reference} #{@tracking ? @tracking : '' } ")
          end
          xml.staticText(attributes)
        end
        @report_element.to_xml(a_node.children.last)
        @box.to_xml(a_node.children.last) unless @box.nil?
        Nokogiri::XML::Builder.with(a_node.children.last) do |xml|
          xml.text_ {
            xml.cdata(@text)
          }
        end
      end

    end # class 'StaticText'

  end # of module 'Vrxml'
end # of module 'Xls'
