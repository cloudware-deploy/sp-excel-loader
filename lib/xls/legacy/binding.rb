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
      # @param hammer       no comments
      #
      def initialize(sheet:, relationship:'lines', hammer: nil)
        super(sheet: sheet, relationship: relationship)
        @parameters = { legacy: nil, map: nil, translated: {}}
        @fields     = { legacy: nil, map: nil, translated: {}}
        @variables  = { legacy: nil, map: nil, translated: {}}
        @hammer     = hammer
      end

      #
      # Collect and translate 'Binding' data.
      #
      def collect ()
        extracted = []
        # parameters
        @parameters[:legacy] = ::Xls::Vrxml::Binding.get_table(named: 'params_def', at: @worksheet, optional: true)
        if nil != @parameters[:legacy]
          @parameters[:map], @parameters[:translated], _extracted = Binding.table_to_array(table: @parameters[:legacy], worksheet: @worksheet, relationship: @relationship, nce: @nce)
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
          _item = { name: item[:value], value: { __origin__: 'layout//auto' } }
          # hammering zone
          _hkt = (item[:type].to_s + 's').to_sym
          _hkn = item[:value].to_sym
          _ho  = false
          if nil != @hammer && nil != @hammer[_hkt] && nil != @hammer[_hkt] && nil != @hammer[_hkt][_hkn]
            _item[:value][:java_class] = @hammer[_hkt][_hkn][:java_class]
            _item[:value][:__origin__] = "external//hammer"
            _ho = true
          end
          case item[:type]
          when :parameter
            if true == _ho || false == @parameters[:translated].include?(item[:value])
              @parameters[:translated][item[:value]] = _item
            end
          when :field
            if true == _ho || false == @fields[:translated].include?(item[:value])
              @fields[:translated][item[:value]] = _item
            end
          when :variable
            if true == _ho || false == @variables[:translated].include?(item[:value])
              @variables[:translated][item[:value]] = _item
            end
          else
            ::Xls::Vrxml::Log.ERROR(msg: "'%s'?" % [ item[:type].to_s ], exception: ArgumentError)
          end
        end
        # SPECIAL CASE i18n_date_format
        inject_i18n_date_format_if_needed()
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
          h = { 'casper.binding': { editable: {} } }
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
          # editable
          if h.include?(:editable)
            h[:'casper.binding'][:editable][:is] = ( 1 == h[:editable] ? true : false )
            h.delete(:editable)
          end
          # widget
          if h.include?(:widget)
            h[:'casper.binding'][:editable][:widget] ||= {}
            h[:'casper.binding'][:editable][:widget][:type] = h[:widget]
            h.delete(:widget)
          end
          if h.include?(:uri)
            h[:'casper.binding'][:editable][:widget] ||= {}
            h[:'casper.binding'][:editable][:widget][:uri] = h[:uri]
            h.delete(:uri)
          end
          if h.include?(:cc_field_id)
            h[:'casper.binding'][:editable][:widget] ||= {}
            h[:'casper.binding'][:editable][:widget][:cc_field_id] = h[:cc_field_id]
            h.delete(:cc_field_id)
          end
          if h.include?(:cc_field_name)
            h[:'casper.binding'][:editable][:widget] ||= {}
            h[:'casper.binding'][:editable][:widget][:cc_field_name] = h[:cc_field_name]
            h.delete(:cc_field_name)
          end
          if h.include?(:cc_field_patch)
            h[:'casper.binding'][:editable][:patch] ||= {}
            h[:'casper.binding'][:editable][:patch][:field] ||= { relationship: 'TODO', name: 'TODO', type: 'TODO'}
            h[:'casper.binding'][:editable][:patch][:field] = h[:cc_field_patch]
            h.delete(:cc_field_patch)
          end
          # 
          if true == ( h[:'casper.binding'][:editable][:is] || false )
            if nil == h[:'casper.binding'][:editable][:field]
              h[:'casper.binding'][:editable][:field] = { name: id, type: h[:java_class] } # relationship = nil -> defaults to $['lines'][index] when it's a field field
              if k.start_with?('$P{')
                h[:'casper.binding'][:editable][:field][:kind] = 'parameter'
              elsif k.start_with?('$F{')
                h[:'casper.binding'][:editable][:field][:kind] = 'field'
              else
                raise "Invalid 'editable' type: must a parameter or a field - not a #{k}"
              end
              end
          end
          #
          translation[id] = { name: id, value: h, updated_at: nil }
        end
        # done
        return map, translation, extracted
      end

      #
      # Lazy workers helper: inject 'i18n_date_format' parameter.
      #
      def inject_i18n_date_format_if_needed()
        has_dates = false
        @parameters[:translated].each do | _k, _v |
          if 'java.util.Date' == _v[:value][:java_class]
            has_dates = true
            break
          end
        end
        if false == has_dates
          @fields[:translated].each do | _k, _v |
            if 'java.util.Date' == _v[:value][:java_class]
              has_dates = true
              break
            end
          end
        end
        if true == has_dates && false == @parameters[:translated].include?('i18n_date_format')
          @parameters[:translated]["$['i18n_date_format']"] = { name: "$['i18n_date_format']", value: { java_class: 'java.lang.String', defaultValueExpression: 'dd/MM/yyyy' } }
        end
      end

    end # of class 'Binding'      

  end # of module 'Legacy'
end # of module 'Xls'