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

    class Group

      attr_accessor :order
      attr_accessor :name
      attr_accessor :group_expression
      attr_accessor :group_header
      attr_accessor :group_footer
      attr_accessor :is_start_new_page
      attr_accessor :is_reprint_header_on_each_page

      def initialize (order:, name: nil)
        @order = order
        @name = name || 'Group1'
        @group_expression  = "$['lines'][index]['data_row_type']"
        @is_start_new_page = nil
        @is_reprint_header_on_each_page = nil
        @group_header = GroupHeader.new(order: -1)
        @group_footer = GroupFooter.new(order: -1)
      end

      def attributes
        rv = Hash.new
        rv['order'] = @order
        rv['name'] = @name
        rv['isStartNewPage'] = @is_start_new_page unless  @is_start_new_page.nil?
        rv['isReprintHeaderOnEachPage'] = @is_reprint_header_on_each_page unless @is_reprint_header_on_each_page.nil?
        return rv
      end

      def to_xml (a_node)
        Nokogiri::XML::Builder.with(a_node) do |xml|
          xml.group(attributes)  {
            unless group_expression.nil?
              xml.groupExpression {
                xml.cdata @group_expression
              }
            end
          }
        end
        @group_header.to_xml(a_node.children.last)
        @group_footer.to_xml(a_node.children.last)
      end

    end
  
  end # of module 'Vrxml'
end # of module 'Xls'
