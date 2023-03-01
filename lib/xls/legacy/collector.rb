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

module Xls
  module Legacy

    class TheCollector

      def initialize(sheet:, relationship:'lines')
        @worksheet    = sheet
        @relationship = relationship
        @nce          = {} # not converted expressions
      end

    end

    class Collector
      
      attr_accessor :bands
      attr_accessor :binding

      #
      # Initialize a 'collector'.
      #
      # @param layout       'Layout' sheet
      # @param binding      'Data Binding' sheet
      # @param relationship for translation purpose.
      #
      def initialize(layout:, binding:, relationship:'lines')
        @bands   = Bands.new(sheet: layout, relationship: relationship)
        @binding = Binding.new(sheet: binding, relationship: relationship)
      end

      #
      # Binding Collector
      #
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
          @parameter    = { legacy: nil, map: nil, translated: {}}
          @fields       = { legacy: nil, map: nil, translated: {}}
          @variables    = { legacy: nil, map: nil, translated: {}}
        end

        #
        # Collect and translate 'Binding' data.
        #
        def collect ()
          # parameters
          @parameter[:legacy] = ::Xls::Vrxml::Binding.get_table(named: 'params_def', at: @worksheet, optional: true)
          if nil != @parameter[:legacy]
            @parameter[:map], @parameter[:translated] = Binding.table_to_array(table: @parameter[:legacy], worksheet: @worksheet, relationship: @relationship, nce: @nce)
          end
          # fields
          @fields[:legacy] = ::Xls::Vrxml::Binding.get_table(named: 'fields_def', at: @worksheet, optional: true)
          if nil != @fields[:legacy]
            @fields[:map], @fields[:translated] = Binding.table_to_array(table: @fields[:legacy], worksheet: @worksheet, relationship: @relationship, nce: @nce)
          end
          # variables          
          @variables[:legacy] = ::Xls::Vrxml::Binding.get_table(named: 'variables_def', at: @worksheet, optional: true)
          if nil != @variables[:legacy]
            @variables[:map], @variables[:translated] = Binding.table_to_array(table: @variables[:legacy], worksheet: @worksheet, relationship: @relationship, nce: @nce, alt_id: :name)
          end
        end

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
          translation = {}
          map.each do | k, v |
            id = ::Xls::Vrxml::Expression.translate(uri: 'TODO', expression: k, relationship: relationship, nc: nce)
            h = {}
            v.each do | k1, v1 |
              next if v1.nil?
              if v1.is_a?(String) && [:name, :expression, :initial_expression].include?(k1)
                h[k1] = ::Xls::Vrxml::Expression.translate(uri: 'TODO', expression: v1, relationship: relationship, nc: nce)
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
          return map, translation
        end

      end # of class 'Binding'

      class Bands < TheCollector

        attr_accessor :map
        attr_accessor :elements

        #
        # Initialize a 'Bands' collector.
        #
        # @param sheet        'Layout' sheet
        # @param relationship for translation purpose.
        #
        def initialize(sheet:, relationship:'lines')
          super(sheet: sheet, relationship: relationship)
          @map         = {}
          @map[:bands] = { legacy: {} }
          @map[:other] = { legacy: { report: {}, group: {}, other:{}, unused: {} } }
          @empty_rows  = [] 
          @cz_comments = []
          @elements    = { legacy: {} }
        end
          
        #
        # Collect and translate 'Bands' data.
        #
        def collect()

          # collect bands
          @band_type = nil
          for row in @worksheet.dimension.ref.row_range
            next if @worksheet[row].nil?
            next if @worksheet[row][0].nil?
            row_tag = map_row_tag(tag: @worksheet[row][0].value.to_s)
            next if row_tag.nil? || 0 == row_tag.length
            if @band_type != row_tag
              process_row_mtag(row: row, row_tag: row_tag)
            end
            if nil != @band_type
              @map[:bands][:legacy][@band_type][:end_row] = row
            end
          end

          # collect bands cells
          @elements[:legacy] = {}
          has_comments = nil != @worksheet.comments && @worksheet.comments.size > 0 && nil != @worksheet.comments[0].comment_list
          @map[:bands][:legacy].each do | name, properties |

            @elements[:legacy][name] = []

            for row in properties[:start_row]..properties[:end_row] do

              column = 1
              r_data = @worksheet[row]
              while column < r_data.size do
                cell = nil
                # has value?
                if nil != r_data[column].value
                  # sanitize
                  value = sanitize(r_data[column].value.strip)
                  # still valid?
                  if 0 != value.length                     
                    # track
                    cell = { hint: RubyXL::Reference.new(row,column).to_s, row: row, column: column, value: value, comments: [] }
                  end                 
                end
                # collect comments
                if nil != cell 
                  if true == has_comments
                    @worksheet.comments[0].comment_list.each_with_index do | comment, index |
                      if ! ( comment.ref.col_range.begin == column && comment.ref.row_range.begin == row )
                        next
                      end
                      comment.text.to_s.lines.each do |text|
                        text.strip!
                        next if text == '' or text.nil?
                        idx = text.index(':')
                        next if idx.nil?
                        tag   = text[0..(idx-1)]
                        value = text[(idx+1)..-1]
                        next if tag.nil? or value.nil?
                        tag.strip!
                        value.strip!
                        cell[:comments] << { hint: RubyXL::Reference.new(row,column).to_s, row: comment.ref.row_range.begin, column: comment.ref.col_range.begin, tag: tag, value: value }
                        @cz_comments << index
                      end
                    end
                  end
                  @elements[:legacy][name] << cell
                end
                # next
                column += 1

              end # while

            end # for

          end # each

          # translate
          translated = {}
          @map.each do | k, h |
            translated[k] = {}
            h[:legacy].each do | k1, v1 |
              t = { name: k1, value: {}, updated_at: Time.now.utc.to_s }
              v1.each do | k2, v2 |
                next if [:start_row, :end_row, :elements].include?(k2)
                if v2.is_a?(String)                  
                  t[:value][k2.to_sym] = ::Xls::Vrxml::Expression.translate(uri: 'TODO', expression: v2, relationship: @relationship, nc: @nce)
                else
                  t[:value][k2.to_sym] = v2
                end
              end
              translated[k][k1] = t
            end
          end
          # elements
          @elements[:translated] = { fields: [], parameters: [], variables: [], cells:[] }
          @elements[:legacy].each do | band, elements |

            elements.each do | element |

              pfv = nil
              exp = nil
              expression = Vrxml::Expression.translate(uri: 'TODO', expression: element[:value], relationship: @relationship, nc: @nce)
              ( Vrxml::Expression.extract(expression: expression) || [] ).each do | e |
                case e[:type]
                when :param
                  pfv ||=[]
                  pfv << { ref: element[:hint], append: :parameters, type: e[:type], name: e[:value] }
                when :field
                  pfv ||=[]
                  pfv << { ref: element[:hint], append: :fields, type: e[:type], name: e[:value] }
                when :variable
                  pfv ||=[]
                  pfv << { ref: element[:hint], append: :variables, type: e[:type], name: e[:value] }
                else
                  raise "???"
                end
              end

              if nil == pfv
                exp = { ref: element[:hint] }
                if ( m = expression.match(/\$SE\{(.*)\}/) )
                  exp[:properties] = [{ name: 'textFieldExpression', value: m[1] } ]
                  # TODO: 2.0 exp[:expression] = "$SE{#{m[1]}}"
                  exp[:expression] = m[1].strip
                else
                  exp[:expression] = expression
                end
              end

              # comments 2 fields or expr
              element[:comments].each do | comment |
                #
                case comment[:tag]
                when 'PT', 'pattern'
                  property = { name: 'pattern', value: Vrxml::Expression.translate(uri: 'TODO', expression: comment[:value], relationship: @relationship, nc: @nce) }
                else
                  puts "tag: #{comment[:tag]}, value: #{comment[:value]}".red
                  next
                end
                #
                if nil != pfv
                  pfv[:properties] ||= []
                  pfv[:properties] << property
                elsif nil != exp
                  exp[:properties] ||= []
                  exp[:properties] << property
                else 
                  raise "WTF?"
                end
              end
              #
                            
              # pfv?
              if nil != pfv
                # add all possible missing parameters / fields / variables
                pfv.each do | _item |
                    _item[:properties] ||= [] 
                    _item[:properties] << { name: 'java_class', value: 'java.lang.String' }
                    @elements[:translated][_item[:append]] << { name: _item[:name], ref: _item[:ref] }
                end
              elsif nil != exp
                exp[:properties] ||= []
                exp[:properties] << { name: 'java_class', value: 'java.lang.String' }
                @elements[:translated][:cells] << exp
              else 
                raise "WTF?"
              end
              #

            end
          end
          #
          translated.each do | k, v |
            @map[k][:translated] = v
          end
          # special handling
          o = @map[:other][:translated].clone
          o.each do | k, v |
            v[:name] = v[:name].to_s.upcase
            @map[:other][k] = v
          end          
        
        end

        #
        # Cleanup 'Bands' legacy data.
        #
        def cleanup()
          # clear comments and empty rows
          @worksheet.change_column_width(0)
          # @empty_rows.each do | row |
          #   ap row
          #   @worksheet.delete_row(row)
          # end
          # require 'byebug' ; debugger
          #
          # 
          @worksheet.comments[0].comment_list.delete_if.with_index { |_, index| @cz_comments.include? index }
        end

      private

        def sanitize(value)
          # try to fix bad expressions
          if value.match(/^[^$"']/) && ( value.include?("$P{") || value.include?("$F{") || value.include?("$V{") || value.include?("$[") || value.include?("$.") || value.include?("$.$$V") )
            _parts = value.split(' ')
            if _parts.count > 1
              _value = ''
              _parts.each do | _part |
                if _part.match(/^[$].*/) || _part.match(/^\(\$.*/)
                  _value += "+ #{_part} "
                else
                  _value += "+ '#{_part} '"
                end
              end
              if _value.length > 2
                _value = _value[2..-1]
              end
              value = _value.strip
              ap value
            end
          end
          value
        end

        def map_row_tag(tag:, allow_sub_bands: true)
          unless allow_sub_bands
            match = tag.match(/\A(TL|SU|BG|PH|CH|DT|CF|PF|LPF|ND)\d*:\z/)
            if match != nil and match.size == 2
              return match[1] + ':'
            end
          end
          tag
        end
    
        def process_row_mtag (row:, row_tag:)
          if row_tag.nil? or row_tag.lines.size == 0
            process_row_tag(row: row, tag: row_tag)
          else
            row_tag.lines.each do |tag|
              process_row_tag(row: row, tag: tag)
            end
          end
        end
  
        def process_row_tag(row:, tag:)
          clear = false
          case tag
          when /BG\d*:/
            @band_type = tag
            @map[:bands][:legacy][tag] ||= { start_row: row, end_row: row }
          when /TL\d*:/
            @band_type = tag
            @map[:bands][:legacy][tag] ||= { start_row: row, end_row: row }
          when /PH\d*:/
            @band_type = tag
            @map[:bands][:legacy][tag] ||= { start_row: row, end_row: row }
          when /CH\d*:/
            @band_type = tag
            @map[:bands][:legacy][tag] ||= { start_row: row, end_row: row }
          when /DT\d*/          
            @band_type = tag
            @map[:bands][:legacy][tag] ||= { start_row: row, end_row: row }
          when /CF\d*:/
            @band_type = tag
            @map[:bands][:legacy][tag] ||= { start_row: row, end_row: row }
          when /PF\d*:/
            @band_type = tag
            @map[:bands][:legacy][tag] ||= { start_row: row, end_row: row }
          when /LPF\d*:/
            @band_type = tag
            @map[:bands][:legacy][tag] ||= { start_row: row, end_row: row }
          when /SU\d*:/
            @band_type = tag
            @map[:bands][:legacy][tag] ||= { start_row: row, end_row: row }
          when /ND\d*:/
            @band_type = tag
            @map[:bands][:legacy][tag] ||= { start_row: row, end_row: row }
          when /GH\d*:/
            @band_type = tag
            @map[:bands][:legacy][tag] ||= { start_row: row, end_row: row }
          when /GF\d*:/
            @band_type = tag
            @map[:bands][:legacy][tag] ||= { start_row: row, end_row: row }
          when /Orientation:.+/i
            @map[:other][:legacy][:other][:orientation] = tag.split(':')[1].strip
            clear = true
          when /Size:.+/i                    
            @map[:other][:legacy][:other][:size] = tag.split(':')[1].strip
            clear = true
          when /VScale:.+/i
            @map[:other][:legacy][:other][:vscale] = tag.split(':')[1].strip.to_f
            clear = true
          when /Report.isTitleStartNewPage:.+/i
            @map[:other][:legacy][:report][:isTitleStartNewPage] = ::Xls::Vrxml::Binding.to_b(tag.split(':')[1].strip)
            clear = true
          when /Report.leftMargin:.+/i
            @map[:other][:legacy][:report][:leftMargin] = tag.split(':')[1].strip.to_i
            clear = true
          when /Report.rightMargin:.+/i
            @map[:other][:legacy][:report][:rightMargin] = tag.split(':')[1].strip.to_i
            clear = true
          when /Report.topMargin:.+/i
            @map[:other][:legacy][:report][:topMargin] = tag.split(':')[1].strip.to_i
            clear = true
          when /Report.bottomMargin:.+/i
            @map[:other][:legacy][:report][:bottomMargin] = tag.split(':')[1].strip.to_i
            clear = true
          when /Group.expression:.+/i
            @map[:other][:legacy][:group][:expression] = tag.split(':')[1]
            clear = true
          when /Group.isStartNewPage:.+/i
            @map[:other][:legacy][:group][:isStartNewPage] = ::Xls::Vrxml::Binding.to_b(tag.split(':')[1].strip)
            clear = true
          when /Group.isReprintHeaderOnEachPage:.+/i
            @map[:other][:legacy][:group][:isReprintHeaderOnEachPage] = ::Xls::Vrxml::Binding.to_b(tag.split(':')[1].strip)
            clear = true
          when /CasperBinding:*/                    # TODO 2.0: ?
            clear = true
          when /BasicExpressions:.+/i               # TODO 2.0: ?
            clear = true
          when /Style:.+/i                          # TODO 2.0: ?
            clear = true
          when /Query:.+/i, /Id:.+/i                # ignored
            clear = true
          when /Band.splitType:.+/i, /IsReport:.+/i # ignored
            clear = true
          else
            @band_type = nil
          end
          # comments
          if nil != @band_type && @worksheet.comments != nil && @worksheet.comments.size > 0 && @worksheet.comments[0].comment_list != nil
            @worksheet.comments[0].comment_list.each_with_index do |comment, index|
              if comment.ref.col_range.begin == 0 && comment.ref.row_range.begin == row
                comment.text.to_s.lines.each do |text|
                  text.strip!
                  next if text == ''
                  tag, value =  text.split(':')
                  next if value.nil? || tag.nil?
                  tag.strip!
                  value.strip!
                  case tag
                  when 'PE' , 'printWhenExpression'
                    if false == @map[:bands][:legacy][@band_type].include?(:printWhenExpression)
                      # TODO 2.0: ? transform_expression(value) # to force declaration of paramters/fields/variables
                      # ap ::Xls::Vrxml::Expression.translate(uri: 'TODO', expression: value, relationship: relationship, nc: nce)
                      @map[:bands][:legacy][@band_type][:printWhenExpression] = value
                      @cz_comments << index
                    end
                  when 'AF', 'autoFloat'
                    @map[:bands][:legacy][@band_type][:auto_float]  = ::Xls::Vrxml::Binding.to_b(value)
                    @cz_comments << index
                  when 'AS' , 'autoStretch'
                    @map[:bands][:legacy][@band_type][:autoStretch] = ::Xls::Vrxml::Binding.to_b(value)
                    @cz_comments << index
                  when 'splitType'
                    @map[:bands][:legacy][@band_type][:splitType] = value
                    @cz_comments << index
                  when 'stretchType'
                    @map[:bands][:legacy][@band_type][:stretchType] = value
                    @cz_comments << index
                  when 'lineParentIdField'
                      # TODO 2.0: edition
                  when 'dataRowTypeAttrName'
                    # TODO 2.0: edition
                  else
                  end
                end # ... lines.each ...
              end
            end
          end
          # clear data
          if true == clear
            @worksheet.add_cell(row, 0, '', nil, true)
            @empty_rows << row
          end
        end

      end # of class 'Bands'

    end # of class 'Collector'

    end # of module 'Legacy'
end # of module 'Xls'