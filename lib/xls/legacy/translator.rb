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
        @layout_sheet = ::Xls::Vrxml::Binding.get_sheet(named: 'Layout', at: @workbook)
        # collect 'legacy' data
        @legacy_binding_sheet = ::Xls::Vrxml::Binding.get_sheet(named: 'Data binding', at: @workbook)
        #
        @band_type = nil
      end
      

      def save(uri:)
        # @comments = load_comments()
        # @comments.each do | k, v |
        # end
        # if @layout_sheet.comments
        #   @layout_sheet.comments.clear
        # end

        #
        # Collect Data
        #
        @collector = Collector.new(layout: @layout_sheet, binding: @legacy_binding_sheet)
        
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
          parameters: ( @collector.binding.parameters || {} ) [:translated],
          fields:     ( @collector.binding.fields     || {} ) [:translated],
          variables:  ( @collector.binding.variables  || {} ) [:translated],
          bands:      ( @collector.bands.map[:bands]  || {} ) [:translated],
          other:      ( @collector.bands.map[:other]  || {} ) [:translated],
        }
                
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
              v.each do | _, f |
                if f.is_a?(Hash)
                  @binding_sheet.add_cell(r , _column, f.to_json)
                else
                  @binding_sheet.add_cell(r , _column, f.to_s)
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

        # validate
        # TODO 2.0: zero comments should remain

        # done
        @workbook.save(uri)
      end

      private

      def load_comments()
        @comments = {}
        if nil == @layout_sheet.comments || 0 == @layout_sheet.comments.size || nil == @layout_sheet.comments[0].comment_list
          return @comments
        end
        @layout_sheet.comments[0].comment_list.each do |comment|
          next if nil == comment.text
          lines = []
          comment.text.to_s.lines.each do |text|
            text.strip!
            next if text == '' or text.nil?
            lines << text
          end
          next if 0 == lines.count
          @comments[comment.ref.to_s] = { lines: lines, obj: comment }
        end
        return @comments
      end
      
    end # of class 'Translator'

  end # of module 'Legacy'
end # of module 'Xls'