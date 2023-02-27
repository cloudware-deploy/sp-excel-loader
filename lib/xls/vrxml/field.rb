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

    class Field

      @@expression = /^\$\['[a-zA-Z0-9_#]+\'\]\[index\]\['[a-zA-Z0-9_#]+'\]$/
      def self.expr
        @@expression
      end

      attr_accessor :name
      attr_accessor :java_class
      attr_accessor :description
      attr_accessor :default_value_expression
      # TODO AG: unused? attr_accessor :is_for_prompting

      attr_accessor :binding

      def initialize(name:, java_class: nil, binding: nil)
        if ! Field.expr().match name
          raise "Invalid 'field' name '#{name}'!"
        end
        @name                     = name
        @binding                  = binding || { __origin__: 'auto' }
        @java_class               = java_class || @binding[:java_class] || 'java.lang.String'
        @description              = @binding[:description] || nil       
        @default_value_expression = @binding[:default]     || @binding[:default_value_expression]
      end

      def attributes
        rv = Hash.new
        rv['name']  = @name
        rv['class'] = @java_class
        return rv
      end

      def to_xml (a_node)
        Nokogiri::XML::Builder.with(a_node) do |xml|
          if nil != @binding[:'__origin__'] && 'auto' == @binding[:'__origin__']
            xml.comment(" Warning: #{self.class.name} named #{@name} type was NOT declared, assuming #{@java_class} ")
          end
          xml.field(attributes) {
            unless @description.nil?
              xml.fieldDescription {
                xml.cdata @description
              }
            end
          }
        end
      end

    end

  end # of module 'Vrxml'
end # of module 'Xls'
