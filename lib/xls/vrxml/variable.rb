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

    class Variable

      @@expression = /^\$\.\$\$VARIABLES\[index\]\['[a-zA-Z0-9_#]+'\]/
      def self.expr
        @@expression
      end
      @@capture = /^\$\.\$\$VARIABLES\[index\]\['([a-zA-Z0-9_#]+)'\]/

      @@known_variables = [
        { name: 'LOCALE'	    	        , java_class: 'java.lang.String' , calculation: 'System' },
        { name: 'RENDERER_ID'	          , java_class: 'java.lang.String' , calculation: 'System' },
        { name: 'CONTINUOUS_PAGE_NUMBER', java_class: 'java.lang.Integer', calculation: 'System' },
        { name: 'PAGE_NUMBER'	          , java_class: 'java.lang.Integer', calculation: 'System' },
        { name: 'ON_ALL_ROWS_PROCESSED' , java_class: 'java.lang.Boolean', calculation: 'System' },
        { name: 'ON_LAST_COPY'		      , java_class: 'java.lang.Boolean', calculation: 'System' },
        { name: 'ON_LAST_PAGE'		      , java_class: 'java.lang.Boolean', calculation: 'System' },
        { name: 'ON_LAST_DOCUMENT'	    , java_class: 'java.lang.Boolean', calculation: 'System' },
        { name: 'SIGNATURE_VISIBLE'	    , java_class: 'java.lang.Boolean', calculation: 'System' },
        { name: 'NUMBER_OF_COPIES'	    , java_class: 'java.lang.Integer', calculation: 'System' },
        { name: 'COPY_NUMBER'		        , java_class: 'java.lang.Integer', calculation: 'System' },
        { name: 'NUMBER_OF_DOCUMENTS'   , java_class: 'java.lang.Integer', calculation: 'System' },
        { name: 'DOCUMENT_NUMBER'	      , java_class: 'java.lang.Integer', calculation: 'System' },
        #
        { name: 'PAGE_COUNT'	          , java_class: 'java.lang.Integer', calculation: 'System' },
        { name: 'REPORT_COUNT'		      , java_class: 'java.lang.Integer', calculation: 'System' },
        { name: 'REMAINING_COUNT'       , java_class: 'java.lang.Integer', calculation: 'System' },
      ]
      def self.known_variables
        @@known_variables
      end
      
      attr_accessor :name
      attr_accessor :java_class
      attr_accessor :calculation
      attr_accessor :reset_type
      attr_accessor :variable_expression
      attr_accessor :initial_value_expression
      attr_accessor :presentation # TODO 2.0 : review usage
      
      attr_accessor :binding

      def initialize (name:, java_class: nil, binding: nil)
        if ! Variable.expr().match name
          raise "Invalid 'variable' name '#{name}'!"
        end
        @name                     = name
        @binding                  = binding || { __origin__: 'auto' }
        @java_class               = java_class || @binding[:java_class] || 'java.lang.String'
        @calculation              = @binding[:calculation]        || 'System'
        @reset_type               = @binding[:reset]              || @binding[:reset_type]
        @variable_expression      = @binding[:expression]         || @binding[:variable_expression]
        @initial_value_expression = @binding[:initial_expression] || @binding[:initial_value_expression]
        @presentation             = @binding[:presentation]
      end

      def attributes
        rv = Hash.new
        rv['name']        = @name.match(@@capture)[1]
        rv['class']       = @java_class
        rv['calculation'] = @calculation
        rv['resetType']   = @reset_type unless @reset_type.nil? or @reset_type == 'None'
        rv['resetGroup']  = 'Group1' if @reset_type == 'Group'
        return rv
      end

      def to_xml (a_node)
        Nokogiri::XML::Builder.with(a_node) do |xml|
          if nil != @binding[:'__origin__'] && 'auto' == @binding[:'__origin__']
            xml.comment(" Warning: #{self.class.name} named #{@name} type was NOT declared, assuming #{@java_class} ")
          end
          xml.variable(attributes) {
            unless @variable_expression.nil?
              xml.variableExpression {
                xml.cdata @variable_expression
              }
            end
            unless @initial_value_expression.nil?
              xml.initialValueExpression {
                xml.cdata @initial_value_expression
              }
            end
          }
        end
      end

    end

  end # of module 'Vrxml'
end # of module 'Xls'
