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

    class Parameter

      @@expression = /^\$\['[a-zA-Z0-9_#]+\'\]$/
      def self.expr
        @@expression
      end
      
      @@capture = /(\$\['[a-zA-Z0-9_#]+\'\])+(?!\[)+/
      def self.capture
        @@capture
      end

      attr_accessor :name
      attr_accessor :java_class
      attr_accessor :description
      attr_accessor :default_value_expression
      attr_accessor :is_for_prompting
      
      attr_accessor :binding
      
      def initialize(name:, java_class: nil, binding: nil)
        if ! Parameter.expr().match name
          if ! Variable.expr().match name
            _name = "$['#{name}']"
            Vrxml::Log.WARNING(msg: "Invalid 'parameter' name '#{name} - assuming #{_name} !")
            @name                   = _name
          else
            @name                   = name
          end
        end
        @name                     = name
        @is_for_prompting         = false
        @binding                  = binding || { __origin__: 'auto' }
        @java_class               = java_class || @binding[:java_class] || 'java.lang.String'
        @description              = @binding[:description] || nil
        @default_value_expression = @binding[:default] || @binding[:defaultValueExpression] || name
      end

      def attributes
        rv = Hash.new
        rv['name']  = @name
        rv['class'] = @java_class
        rv['isForPrompting'] = false if @is_for_prompting == false
        return rv
      end

      def to_xml (a_node)
        Nokogiri::XML::Builder.with(a_node) do |xml|
          if nil != @binding[:'__origin__'] && 'auto' == @binding[:'__origin__']
            xml.comment(" Warning: #{self.class.name} named #{@name} type was NOT declared, assuming #{@java_class} ")
          end
          xml.parameter(attributes) {
            unless @description.nil?
              xml.parameterDescription {
                xml.cdata @description
              }
            end
            unless @default_value_expression.nil?
              xml.defaultValueExpression {
                xml.cdata @default_value_expression
              }
            end
          }
        end
      end

    end # class 'Parameter'

  end # of module 'Vrxml'
end # of module 'Xls'
