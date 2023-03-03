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

require_relative 'object'

module Xls
  module Legacy

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
        @elements    = { legacy: {}, translated: { fields: [], parameters: [], variables: [], cells:[] } }
      end
          
      #
      # Collect and translate 'Bands' data.
      #
      def collect()

        # collect bands
        @band_type = nil
        for row in @worksheet.dimension.ref.row_range
          if @worksheet[row].nil? || @worksheet[row][0].nil?
            next
          end
          row_tag = map_row_tag(tag: @worksheet[row][0].value.to_s)
          if row_tag.nil? || 0 == row_tag.length
            next
          end
          if @band_type != row_tag
            process_row_mtag(row: row, row_tag: row_tag)
          end
          if nil != @band_type
            @map[:bands][:legacy][@band_type][:end_row] = row
          end
        end # for row

        #
        has_comments = nil != @worksheet.comments && @worksheet.comments.size > 0 && nil != @worksheet.comments[0].comment_list

        # collect bands cells
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
                      end # lines.each 
                  end # each_with_index
                end # if has_comments
                  @elements[:legacy][name] << cell
              end # if nil != cells
              # next
              column += 1

            end # while column 

          end # for row 

        end # @map[:bands][:legacy].each

        # translate
        translated = {}
        @map.each do | k, h |
          translated[k] = {}
          h[:legacy].each do | k1, v1 |
            t = { name: k1, value: {}, updated_at: Time.now.utc.to_s }
            v1.each do | k2, v2 |
              if [:start_row, :end_row, :elements].include?(k2)
                next
              end
              if v2.is_a?(String)
                _exp, _ext = Vrxml::Expression.translate(expression: v2, relationship: @relationship, nce: @nce)
                if _ext.count > 0
                  ::Xls::Vrxml::Log.TODO(msg: "@ #{__FILE__}:#{__LINE__} - #{__method__} : Add possible MISSING parameter(s)/field(s)/variable(s) %d" % [ _ext.count])
                end
                t[:value][k2.to_sym] = _exp
              else
                t[:value][k2.to_sym] = v2
              end
            end # v1.each
            translated[k][k1] = t
          end # h[:legacy].each
        end # @map.each

        # elements
        @elements[:legacy].each do | band, elements |
          elements.each do | element |

            pfv = nil
            exp = nil
            expression, _extracted = Vrxml::Expression.translate(expression: element[:value], relationship: @relationship, nce: @nce)
            _extracted.each do | e |
              case e[:type]
              when :parameter
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
            end # each

            #
            if nil == pfv || pfv.count > 1
              pfv = nil
              exp = { ref: element[:hint] }
              if ( m = expression.match(/\$SE\{(.*)\}/) )
                exp[:properties] = [{ name: 'textFieldExpression', value: m[1] } ]
                exp[:expression] = m[1].strip
              else
                exp[:expression] = expression
              end
            else
              if ( m = expression.match(/\$SE\{(.*)\}/) )
                pfv[0][:properties] ||= []
                pfv[0][:properties] = [{ name: 'textFieldExpression', value: m[1] } ]
              end
            end

            # comments 2 fields or expr
            element[:comments].each do | comment |
              #
              property = nil
              case comment[:tag]
              when 'PT', 'pattern'
                _exp, _ext = Vrxml::Expression.translate(expression: comment[:value], relationship: @relationship, nce: @nce)
                if _ext.count > 0
                  ::Xls::Vrxml::Log.TODO(msg: "@ #{__FILE__}:#{__LINE__} - #{__method__} : Add possible MISSING parameter(s)/field(s)/variable(s) %d" % [ _ext.count])
                end
                property = { name: 'pattern', value: _exp }
              else
                puts "tag: #{comment[:tag]}, value: #{comment[:value]}".red
                next
              end 
              # case
              if nil != pfv
                if pfv.count > 1
                  # it should be already resolved as an expression ( see code above )
                  raise "WTF?"
                else
                  pfv[0][:properties] ||= []
                  pfv[0][:properties] << property
                end
              elsif nil != exp
                exp[:properties] ||= []
                exp[:properties] << property
              else 
                raise "WTF?"
              end # if
            end # each

            # pfv?
            if nil != pfv
              # add all possible missing parameters / fields / variables
              pfv.each do | _item |
                _item[:properties] ||= [] 
                _item[:properties] << { name: 'java_class', value: 'java.lang.String' }
                @elements[:translated][_item[:append]] << { name: _item[:name], ref: _item[:ref] }
              end # pfv.each
            elsif nil != exp
              exp[:properties] ||= []
              exp[:properties] << { name: 'java_class', value: 'java.lang.String' }
              @elements[:translated][:cells] << exp
            else 
              raise "WTF?"
            end # if

          end # elements.each
        end #  @elements[:legacy].each
        #
        translated.each do | k, v |
          @map[k][:translated] = v
        end # translated.each
        # special handling
        o = @map[:other][:translated].clone
        o.each do | k, v |
          v[:name] = v[:name].to_s.upcase
          @map[:other][k] = v
        end # o.each

      end # elements.each

      #
      # Cleanup 'Bands' legacy data.
      #
      def cleanup()
        @worksheet.change_column_width(0)
        @worksheet.comments[0].comment_list.delete_if.with_index { |_, index| @cz_comments.include? index }
      end

      private     

      #
      # Sanitize a cell value.
      #
      # value Cell value to sanitize.
      #
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
          end # _parts.count > 1
        end
        # done
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
      end # map_row_tag

      def process_row_mtag(row:, row_tag:)
        if row_tag.nil? or row_tag.lines.size == 0
          process_row_tag(row: row, tag: row_tag)
        else
          row_tag.lines.each do |tag|
              process_row_tag(row: row, tag: tag)
          end
        end
      end # process_row_mtag

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
                      _exp, _ext = Vrxml::Expression.translate(expression: value, relationship: @relationship, nce: @nce)
                      if _ext.count > 0
                        ::Xls::Vrxml::Log.TODO(msg: "@ #{__FILE__}:#{__LINE__} - #{__method__} : Add possible MISSING parameter(s)/field(s)/variable(s) %d" % [ _ext.count])
                      end
                      @map[:bands][:legacy][@band_type][:printWhenExpression] = _exp
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
                else
                  ::Xls::Vrxml::Log.WHAT_IS(msg: "@ #{__FILE__}:#{__LINE__} - #{__method__} : TAG #{tag}")
                end # case
              end # ... lines.each ...
            end # if
          end # each_with_index
        end # if
        # clear data
        if true == clear
          @worksheet.add_cell(row, 0, '', nil, true)
          @empty_rows << row
        end
      end # process_row_tag

    end # of class 'Bands'

  end # of module 'Legacy'
end # of module 'Xls'