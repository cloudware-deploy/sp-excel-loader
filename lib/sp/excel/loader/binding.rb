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
# encoding: utf-8
#

module Sp
  module Excel
    module Loader

      class Binding

        attr_accessor :sheet
        attr_accessor :table
        attr_accessor :columns
        attr_accessor :map

        #
        # Initializer binding for a workbook
        #
        # @param @workbook Pre-loaded workbook.
        #
        def initialize(workbook:)
          @workbook = workbook
          @columns  = {}
        end

        #
        # Load 'binding' data from the workbook.
        #
        def load()
          @sheet = Binding.get_sheet(named: 'Binding', at: @workbook)
          @table = Binding.get_table(named: 'CASPER', at: @sheet)
          @columns['Name']       = get_table_column_index(table: table, named: 'Name')
          @columns['Value']      = get_table_column_index(table: table, named: 'Value')
          @columns['Updated At'] = get_table_column_index(table: table, named: 'Updated At')
          @map = {}
          Binding.iterate_table(table: @table, at: @sheet) do | row, cells |
            data = {}
            @columns.each do | k, _ |
              data[k] = {
                cell: cells[@columns[k]],
                name: k,
                value: cells[@columns[k]] ? cells[@columns[k]].value : nil,
                row: row,
                column: @columns[k]
              }
            end
            @map[data['Name'][:value]] = data
          end
          @named_cells = Binding.get_named_cells_map(sheet: Binding.get_sheet(named: 'Layout', at: @workbook), at: @workbook)
        end

        #
        #
        # Patch a field binding data.
        #
        # @param field Field name ( string ) to patch.
        # @param value Field value ( JSON object ) to set.
        #
        def patch(field:, value:)
          values = { 'Value': value, 'Updated At': Time.now }
          # grab field data
          data = @map[field]
          if nil == data
            values['Name'] = field

            row = @sheet.add_row()
            @table.auto_filter.ref = RubyXL::Reference.new(
              @table.auto_filter.ref.row_range.min, @table.auto_filter.ref.row_range.max + 1,
              @table.auto_filter.ref.col_range.min, @table.auto_filter.ref.col_range.max
            )
            @table.ref = "A1:C4" # TODO
            data = {}
            @columns.each do | k, _ |
              data[k] = {
                cell: nil,
                name: k,
                value: nil,
                row: @map.size + 1,
                column: @columns[k]
              }
            end
            @map[data['Name'][:value]] = data
            # raise "data for binding field #{field} NOT found!"
          end
          #
          values.each do | k, v |
            cell = data[k.to_s][:cell]
            # ... update value ...
            ap k.to_s
            ap v
            if nil == cell
              # ... by adding a new cell ...
              sheet.add_cell(data[k.to_s][:row], data[k.to_s][:column], v)
            else
              # ... by replacing cell ...
              sheet.delete_cell(cell.row, cell.column)
              sheet.add_cell(cell.row, cell.column, v)
            end
          end
        end
        
        private
        
        #
        # Get a sheet from a workbook.
        #
        # @param named: Sheet name.
        # @param at: Workbook
        #
        def self.get_sheet(named:, at:)
          at.worksheets.each do |ws|
            if ws.sheet_name == named
              return ws
            end
          end
          raise "Sheet #{named} NOT found!"
        end

        #
        # Obtain a map of named cells to cells reference.
        #
        # @param sheet Sheet where to look for named cells.
        # @param at    Workbook.
        #
        # TODO: this duplicates read_named_cells
        #
        def self.get_named_cells_map(sheet:, at:)
          map = Hash.new
          ref_regexp = sheet.sheet_name + '!\$*([A-Z]+)\$*(\d+)'
          at.defined_names.each do |dn|
            next unless dn.local_sheet_id.nil?
            match = dn.reference.match(ref_regexp)
            if match and match.size == 3
              matched_name = match[1].to_s + match[2].to_s
              if map[matched_name]
                raise "**** Fatal error:\n     duplicate cellname for #{matched_name}: #{@map[matched_name]} and #{dn.name}"
              end
              map[dn.name] = matched_name
            end
          end
          map
        end

        #
        # Get a table from a sheet.
        #
        # @param named Table name.
        # @param at Sheet.
        #
        def self.get_table(named:, at:)
          klass = TableRow.factory named
          at.generic_storage.each do |tbl|
            return tbl if tbl.is_a? RubyXL::Table and tbl.name == named
          end
          raise "Table #{named} NOT found!"
        end

        #
        # Obtain a table column index.
        #
        # @param table Table name.
        # @param named Column name.
        #
        def get_table_column_index(table:, named:)
          table.table_columns.each_with_index do | column, index |
            return index if column.name == named
          end
          raise "Table column #{named} NOT found!"
        end

        #
        # Iterate a table.
        #
        # @param table Table object to iterate.
        # @param at    Sheet that owns table.
        #
        def self.iterate_table(table:, at:)
          ref = RubyXL::Reference.new(table.ref)
          for row in ref.row_range.begin()+1..ref.row_range.end()
            row_cells = []
            ref.col_range.each do |column|
              row_cells << at[row][column]
            end
            yield row, row_cells
          end
        end

      end # of class 'Binding'

    end
  end
end
  