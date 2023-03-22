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

      attr_reader :map
      attr_reader :elements
      attr_reader :named_cells

      #
      # Initialize a 'Bands' collector.
      #
      # @param sheet        'Layout' sheet
      # @param relationship for translation purpose.
      # @param hammer       no comments
      #
      def initialize(sheet:, relationship:'lines', hammer: nil)
        super(sheet: sheet, relationship: relationship)
        @map         = {}
        @map[:bands] = { legacy: {} }
        @map[:other] = { legacy: { report: {}, group: {}, other:{}, unused: {} } }
        @empty_rows  = [] 
        @cz_comments = []
        @elements    = { legacy: {}, translated: { parameters: [], fields: [], variables: [], cells:[] } }
        @auto_naming = { parameters: {}, fields: {}, variables: {}, expressions:{} }
        @named_cells = {}
        @hammer      = hammer
      end
          
      #
      # Collect and translate 'Bands' data.
      #
      def collect(hammer: nil)

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
        has_comments = nil != @worksheet.comments && @worksheet.comments.size > 0 && nil != @worksheet.comments[0].comment_list && @worksheet.comments[0].comment_list.count > 0

        # collect bands cells
        @map[:bands][:legacy].each do | name, properties |

          @elements[:legacy][name] = []

          for row in properties[:start_row]..properties[:end_row] do

            column = 1
            r_data = @worksheet[row]
            while column < r_data.size do
              cell = nil
              # has value?
              if nil != r_data[column] && nil != r_data[column].value
                if r_data[column].value.is_a?(String)
                  value = r_data[column].value.strip
                else
                  value = r_data[column].value
                end
                # still valid?
                if nil != value && ( false == value.is_a?(String) || 0 != value.length )
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

        # elements
        @elements[:legacy].each do | band, elements |
          
          elements.each do | element |

            pfv = nil
            exp = nil

            expression = element[:value]

            tfe = false
            ie = false
            if true == expression.is_a?(String)
              if ( m = expression.match(/\$SE\{(.*)\}/) )
                expression = m[1]
                tfe = true
              elsif ( m = expression.match(/\$RB\{(\$[PFV]{1}\{[a-zA-Z0-9_\-\?]+\})\s*,\s*(\d{1,})\s*,\s*(\d{1,})\}/) )
                # check box: $RB{<field_name>,<unchecked>,<checked>}
                # ( at this point we still need a "JAVA" expression to be translated later on )
                expression = "IF ( ( null == #{m[1]} || #{m[2]} == #{m[1]} ) ; \" \" ; IF ( #{m[3]} == #{m[1]} ; \"X\" ; \" \" ) )"
                tfe = true
              elsif( m = expression.match(/\$CB\{(\$[PFV]{1}\{[a-zA-Z0-9_\-\?]+\})\s*,\s*(\d{1,})\s*,\s*(\d{1,})\}/) )
                # check box: $CB{<field_name>,<unchecked>,<checked>}
                # ( at this point we still need a "JAVA" expression to be translated later on )
                expression = "IF ( ( null == #{m[1]} || #{m[2]} == #{m[1]} ) ; \" \" ; IF ( #{m[3]} == #{m[1]} ; \"X\" ; \" \" ) )"
              elsif ( m = expression.match(/\$CB\{(\$[PFV]{1}\{[a-zA-Z0-9_\-\?]+\})\s*,\s*(\btrue\b|\bfalse\b{1})\s*,\s*(\btrue\b|\bfalse\b{1})\}/) )
                # check box: $CB{<field_name>,<unchecked>,<checked>}
                # ( at this point we still need a "JAVA" expression to be translated later on )
                expression = "IF ( ( null == #{m[1]} || #{m[2]} == #{m[1]} ) ; \" \" ; IF ( #{m[3]} == #{m[1]} ; \"X\" ; \" \" ) )"
              elsif ( m = expression.match(/\$I\{(.*)\}/) )
                expression = m[1]
                ie = true
              end
            end

            if true == expression.is_a?(String)
              expression, _extracted = Vrxml::Expression.translate(expression: expression, relationship: @relationship, nce: @nce)
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
            end

            #
            if nil == pfv || pfv.count > 1
              pfv = nil
              exp = { ref: element[:hint] }
              if true == tfe
                exp[:properties] = [{ name: 'textFieldExpression', value: expression } ]
              elsif true == ie
                exp[:properties] = [{ name: 'imageExpression', value: expression } ]
              else
                exp[:properties] = [{ name: 'text', value: expression } ]
              end
              @auto_naming[:expressions][band] ||= {}
              exp[:__cell__] = { ref: exp[:ref] , value: expression, name: "#{band.to_s.gsub(':', '').upcase}_EXPRESSION_#{@auto_naming[:expressions][band].count + 1}" }
              @auto_naming[:expressions][band][exp[:__cell__][:ref]] = exp[:__cell__][:name]
            else
              pfv[0][:properties] ||= []
              if true == tfe
                pfv[0][:properties] = [{ name: 'textFieldExpression', value: expression } ]
              elsif true == ie
                pfv[0][:properties] = [{ name: 'imageExpression', value: expression } ]
              else
                pfv[0][:properties] = [{ name: 'text', value: expression } ]
              end
              pfv_pkey = ( pfv[0][:type].to_s + 's' ).to_sym
              @auto_naming[pfv_pkey][band] ||= {}
              pfv[0][:__cell__] = { ref: pfv[0][:ref] , value: expression, name: "#{band.to_s.gsub(':', '').upcase}_#{pfv[0][:type].to_s.upcase}_#{@auto_naming[pfv_pkey][band].count}" }
              @auto_naming[pfv_pkey][band][pfv[0][:__cell__][:ref]] = pfv[0][:__cell__][:name]
            end

            # comments 2 fields or expr
            element[:comments].each do | comment |
              #
              property = nil
              case comment[:tag]
              when 'PT', 'pattern'
                _exp, _ = Vrxml::Expression.translate(expression: comment[:value], relationship: @relationship, nce: @nce)
                property = { name: 'pattern', value: _exp }
              when 'AS' , 'autoStretch'
                _exp, _ = Vrxml::Expression.translate(expression: comment[:value], relationship: @relationship, nce: @nce)
                property = { name: 'autoStretch', value: _exp }
              when 'PE' , 'printWhenExpression'
                _exp, _ext = Vrxml::Expression.translate(expression: comment[:value], relationship: @relationship, nce: @nce)
                if _ext.count > 0
                  _ext.each do | item |
                    add_pfv_if_missing(type: item[:type], ref: RubyXL::Reference.new(comment[:row], comment[:column]).to_s, name: item[:value])
                  end
                end
                property = { name: 'printWhenExpression', value: _exp }
              when 'ET', 'evaluationTime'
                property = { name: 'evaluationTime', value: _exp }
              when 'BN', 'blankIfNull' 
                property = { name: 'isBlankWhenNull', value: _exp }
              else
                # log
                ::Xls::Vrxml::Log.TODO(msg: "@ #{__method__}: process tag %s - %s" % [comment[:tag], comment[:value]])
                # next
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

            # ... A.S. not A.I. ...
            _suspicious = ( nil != _extracted && _extracted.count > 1 && false == ['==', '===', '!=', '>' , '<', '!='].any? { |word| expression.include?(word) } && false == expression.start_with?('`') && false == expression.end_with?('`') )

            # pfv?
            if nil != pfv
              # add all possible missing parameters / fields / variables
              pfv.each do | _item |
                _item[:properties] ||= [] 
                _item[:properties] << { name: '__original_java_expression__', value: element[:value]      }
                _item[:properties] << { name: '__composed__'                , value: true            } if nil != _extracted && _extracted.count > 1
                _item[:properties] << { name: '__suspicious__'              , value: _suspicious     } if true == _suspicious                  
                add_pfv_if_missing(type: _item[:append], ref: _item[:ref], name: _item[:name])
                @elements[:translated][:cells] << _item
              end # pfv.each
            elsif nil != exp
              exp[:properties] ||= []
              exp[:properties] << { name: '__original_java_expression__', value: element[:value] }
              exp[:properties] << { name: '__composed__'                , value: true            } if nil != _extracted && _extracted.count > 1
              exp[:properties] << { name: '__suspicious__'              , value: _suspicious     } if true == _suspicious
              @elements[:translated][:cells] << exp
            else 
              raise "WTF?"
            end # if

          end # elements.each
        end #  @elements[:legacy].each
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
                  _ext.each do | item |
                    add_pfv_if_missing(type: item[:type], ref: nil, name: item[:value])
                  end
                end
                t[:value][k2.to_sym] = _exp
              else
                t[:value][k2.to_sym] = v2
              end
            end # v1.each
            translated[k][k1] = t
          end # h[:legacy].each
        end # @map.each        
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

        # @auto_naming -> @named_cells
        @auto_naming.each do | type, values |
          values.each do | band, map |
            map.each do | ref, name |
              @named_cells[ref] = name
            end
          end
        end
      end # collect

      #
      # Cleanup 'Bands' legacy data.
      #
      def cleanup()
        @worksheet.change_column_width(0)
        if @worksheet.comments.size > 0 && nil != @worksheet.comments[0].comment_list
          @worksheet.comments[0].comment_list.delete_if.with_index { |_, index| @cz_comments.include? index }
        end
      end

      private

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
                        _ext.each do | item |
                          add_pfv_if_missing(type: item[:type], ref: nil, name: item[:value])
                        end
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

      #
      # Add a parameter/field/variable if missing.
      #
      # @param type One of parameter/field/variable.
      # @param ref  Cell reference.
      # @paeam name parameter/field/variable name.
      #
      def add_pfv_if_missing(type:, ref:, name:)
        # ... translation ...
        case type
        when :parameter, :parameters
          _type = :parameters
        when :field, :fields
          _type = :fields
        when :variable, :variables
          _type = :variables
          if true == Xls::Vrxml::Variable.is_known_variable(name)
            return
          end
        else
          ::Xls::Vrxml::Log.ERROR(msg: "'%s'?" % [ type.to_s ], exception: ArgumentError)
        end
        # ... do NOT allow duplication ...
        add = true
        (@elements[:translated][_type] || []).each do | e |
          if e[:value] == name
            add = false
            break
          end
        end
        # ...
        java_class = nil
        if nil != @hammer && nil != @hammer[_type] && nil != @hammer[_type][name.to_sym]
          java_class = @hammer[_type][name.to_sym][:java_class]
        end
        # ... add?
        if true == add
          @elements[:translated][_type] << { name: name, __origin__: __method__, ref: ref, java_class: java_class }
        end
      end # add_if_missing

    end # of class 'Bands'

  end # of module 'Legacy'
end # of module 'Xls'