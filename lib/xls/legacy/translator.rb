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

# require 'xls/loader'
# require_relative '../loader/workbookloader'
# require_relative '../vrxml/binding'

module Xls
  module Legacy

    class Translator < ::Xls::Loader::WorkbookLoader      

      FONT_NAME = 'Times New Roman'

      public

      def initialize(uri:)
        super(uri)
        # ... [B] consistency is not a requirement ... ...
        @layout_sheet_name = Xls::Vrxml::Object.guess_layout_sheet(workbook: @workbook)
        # ... [E] consistency is not a requirement ... ...        
        @layout_sheet = ::Xls::Vrxml::Binding.get_sheet(named: @layout_sheet_name, at: @workbook)
        # collect 'legacy' data
        @legacy_binding_sheet = ::Xls::Vrxml::Binding.get_sheet(named: ['Data binding', 'Databinding'], at: @workbook)
        #
        @band_type = nil
        #
        @hammer     = nil
        @hammer_uri = uri + ".json"
        if File.exist?(@hammer_uri)
          @hammer = JSON.parse(File.read(@hammer_uri), symbolize_names: true)
        end
      end
      

      def translate(to:)
        #
        # Collect Data
        #
        @collector = Collector.new(layout: @layout_sheet, binding: @legacy_binding_sheet, hammer: @hammer)
        
        @collector.bands.collect()
        @collector.bands.cleanup()

        @collector.binding.collect()


        # remove 'legacy' binding
        @workbook.worksheets.each_with_index do | sheet, index |
          if @legacy_binding_sheet.sheet_name == sheet.sheet_name 
            @workbook.worksheets.delete_at(index)
            break
          end
        end
        
        #
        # Define new 'translated' tables for new binding sheet.
        #
        tables = {
          parameters:  ( ( @collector.binding.parameters || {} ) [:translated] ).clone,
          fields:      ( ( @collector.binding.fields     || {} ) [:translated] ).clone,
          variables:   ( ( @collector.binding.variables  || {} ) [:translated] ).clone,
          bands:       ( ( @collector.bands.map[:bands]  || {} ) [:translated] ).clone,
          other:       ( ( @collector.bands.map[:other]  || {} ) [:translated] ).clone,
          named_cells: {}
        }

        [:parameters, :fields, :variables ].each do | type |
          if ! @collector.bands.elements[:translated][type]
            next
          end
          tables[type] ||= {}
          @collector.bands.elements[:translated][type].each do | _item |
            if true == tables[type].include?(_item[:name])
              next
            end
            tables[type][_item[:name]] = { name: _item[:name], value: { '__origin__': ( _item[:__origin__] || "\"#{__method__}\"" ), java_class: 'java.lang.String' } }
            if nil != _item[:ref]
              i = RubyXL::Reference.ref2ind(_item[:ref])
              @layout_sheet[i[0]][i[1]].change_contents(_item[:name])
            end
          end
        end

        @collector.bands.elements[:translated][:cells].each do | cell |
          #
          _ref = cell[:__cell__][:ref]
          i = RubyXL::Reference.ref2ind(_ref)
          # go to specific cell and patch it!
          @layout_sheet[i[0]][i[1]].change_contents(cell[:__cell__][:value])
          # add named cell binding info
          value = { ref: _ref, value: {}}
          if nil != cell[:properties]
            cell[:properties].each do | property |
              value[:value][property[:name].to_sym] = property[:value]
            end
          end
          if @collector.bands.named_cells.include?(_ref)
            value[:name] = @collector.bands.named_cells[_ref]
            # inject missing properties
            if nil != tables[cell[:append]] && nil != tables[cell[:append]][cell[:name]] && nil != tables[cell[:append]][cell[:name]][:value]
              tables[cell[:append]][cell[:name]][:value].each do | k, v |
                if false == value[:value].include?(k)                  
                  value[:value][k] = v
                end
              end
            end
          end
          value[:updated_at] ||= Time.now.utc.to_s
          # done
          tables[:named_cells][_ref] = value
        end

        # set named cells
        @collector.bands.named_cells.each do | ref, name |
          @workbook.define_new_name(name, @layout_sheet.ref2abs(ref))
        end

        # add 'Binding' sheet
        @binding_sheet = @workbook.add_worksheet('Binding')

        table_defs = []

        # 
        r = 0
        ::Xls::Vrxml::Binding.all().each do | definition |
          #
          table_def = { name: definition[:name].to_s }
          # separator
          @binding_sheet.add_cell(r, 0, definition[:name])
          @binding_sheet.merge_cells(r, 0, r, 2)
          @binding_sheet.sheet_data[r][0].change_horizontal_alignment('center')
          @binding_sheet.change_column_width(0, 50)
          @binding_sheet.change_column_width(1, 240)
          @binding_sheet.change_column_width(2, 30)
          @binding_sheet.sheet_data[r][0].change_font_bold(true)
          @binding_sheet.sheet_data[r][0].change_font_color('ffffff')
          @binding_sheet.sheet_data[r][0].change_fill('2f61ba')
          @binding_sheet.sheet_data[r][0].change_font_name(FONT_NAME)
          r += 1
          #
          table_def[:start_row]    = r + 1
          table_def[:start_column] = 'A'
          # columns names
          c = 0
          ::Xls::Vrxml::Binding.columns().each do | column |
            @binding_sheet.add_cell(r, c, column)
            @binding_sheet.sheet_data[r][c].change_font_bold(true)
            @binding_sheet.sheet_data[r][c].change_font_name(FONT_NAME)
            c+=1
          end 
          # columns values
          if tables[definition[:key]]
            tables[definition[:key]].each do | _, v |
              r += 1
              _column = 0
              [:name, :value, :updated_at].each do | _c |
                if v[_c].is_a?(Hash)
                  @binding_sheet.add_cell(r , _column, v[_c].to_json)
                else
                  @binding_sheet.add_cell(r , _column, v[_c].to_s)
                end
                _column += 1
              end
            end
          end
          # seal
          table_def[:end_row]    = r + 2
          table_def[:end_column] = 'C'
          table_def[:ref]        = "#{table_def[:start_column]}#{table_def[:start_row]}:#{table_def[:end_column]}#{table_def[:end_row]}"
          table_defs << table_def
          # next
          r += 3
        end

        # define tables
        table_defs.each_with_index do | definition, index |
          @binding_sheet.add_table(id: index, name: definition[:name], ref: definition[:ref], columns: [
            { id: "1", name: "Name" }, { id: "2", name: "Value" }, { id: "3", name: "Updated At" }
          ])  
        end

        # TODO 2.0: check if we have zero missing translations

        # done
        @workbook.save(to)
      end
      
    end # of class 'Translator'

  end # of module 'Legacy'
end # of module 'Xls'