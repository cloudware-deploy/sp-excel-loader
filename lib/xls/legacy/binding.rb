# encoding: utf-8
#
# Copyright (c) 2011-2023 Cloudware S.A. All rights reserved.
#
# This file is part of xls2vrxml.
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

require_relative 'bands'

module Xls
  module Legacy

    class Binding < TheCollector

      attr_accessor :parameters
      attr_accessor :fields
      attr_accessor :variables

      #
      # Initialize a 'Binding' collector.
      #
      # @param sheet        'Data Binding' sheet
      # @param relationship for translation purpose.
      #
      def initialize(sheet:, relationship:'lines')
        super(sheet: sheet, relationship: relationship)
        @parameter = { legacy: nil, map: nil, translated: {}}
        @fields    = { legacy: nil, map: nil, translated: {}}
        @variables = { legacy: nil, map: nil, translated: {}}
      end

      #
      # Collect and translate 'Binding' data.
      #
      def collect ()
        extracted = []
        # parameters
        @parameter[:legacy] = ::Xls::Vrxml::Binding.get_table(named: 'params_def', at: @worksheet, optional: true)
        if nil != @parameter[:legacy]
          @parameter[:map], @parameter[:translated], _extracted = Binding.table_to_array(table: @parameter[:legacy], worksheet: @worksheet, relationship: @relationship, nce: @nce)
          extracted.concat(_extracted)
        end
        # fields
        @fields[:legacy] = ::Xls::Vrxml::Binding.get_table(named: 'fields_def', at: @worksheet, optional: true)
        if nil != @fields[:legacy]
          @fields[:map], @fields[:translated], _extracted = Binding.table_to_array(table: @fields[:legacy], worksheet: @worksheet, relationship: @relationship, nce: @nce)
          extracted.concat(_extracted)
        end
        # variables          
        @variables[:legacy] = ::Xls::Vrxml::Binding.get_table(named: 'variables_def', at: @worksheet, optional: true)
        if nil != @variables[:legacy]
          @variables[:map], @variables[:translated], _extracted = Binding.table_to_array(table: @variables[:legacy], worksheet: @worksheet, relationship: @relationship, nce: @nce, alt_id: :name)
          extracted.concat(_extracted)
        end
        # 
        extracted.each do | item |          
          case item[:type]
          when :parameter
            if false == @parameter[:translated].include?(item[:value])
              @parameter[:translated][item[:value]] = { name: item[:value], value: { __origin__: 'layout//auto' } }
            end
          when :field
            if false == @fields[:translated].include?(item[:value])
              @fields[:translated][item[:value]] = { name: item[:value], value: { __origin__: 'layout//auto' } }
            end
          when :variable
            if false == @variables[:translated].include?(item[:value])
              @variables[:translated][item[:value]] = { name: item[:value], value: { __origin__: 'layout//auto' } }
            end
          else
            ::Xls::Vrxml::Log.ERROR(msg: "'%s'?" % [ item[:type].to_s ], exception: ArgumentError)
          end
        end
      end # collect()

      private

      #
      #
      #
      def self.table_columns(table:)
        map = {}
        table.table_columns.each_with_index do | column , index |
          map[index] = column.name
        end
        map
      end

      #
      #
      #
      def self.table_to_array(table:, worksheet:, relationship:, nce:, alt_id: :id)
        #
        columns = {}
        table.table_columns.each_with_index do | column , index |
          columns[index] = column.name
        end
        #
        map = {}
        ::Xls::Vrxml::Binding.iterate_table(table: table, at: worksheet) do | row, cells |
          j = {}
          cells.each do | cell |
            next if nil == cell 
            j[columns[cell.column].to_sym] = cell.value
          end
          if nil == j[:id] 
            if nil == j[alt_id]
              next
            else
              j[:id] = j[alt_id]
            end
          end
          map[j[:id]] = j
        end
        # translate
        extracted   = [] # parameters / fields / variables from expressions
        translation = {}
        map.each do | k, v |
          id, _ext = Vrxml::Expression.translate(expression: k, relationship: relationship, nce: nce)
          if _ext.count > 0
            extracted.concat(_ext)
          end
          h = {}
          v.each do | k1, v1 |
            next if v1.nil?
            if v1.is_a?(String) && [:name, :expression, :initial_expression].include?(k1)
              h[k1], _ext = Vrxml::Expression.translate(expression: v1, relationship: relationship, nce: nce)
              if _ext.count > 0
                extracted.concat(_ext)
              end
            else
              h[k1] = v1
            end              
          end
          h.delete(:id)
          if h.include?(:editable)
            h[:editable] = ( 1 == h[:editable] ? true : false )
          end
          translation[id] = { name: id, value: h, updated_at: Time.now.utc.to_s }
        end
        # done
        return map, translation, extracted
      end

    end # of class 'Binding'      

  end # of module 'Legacy'
end # of module 'Xls'