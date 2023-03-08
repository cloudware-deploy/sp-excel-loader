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

require 'set'

require_relative 'parameter'
require_relative 'field'
require_relative 'variable'

module Xls
  module Vrxml

    class Converter < Loader::WorkbookLoader

      @@CT_IndexedColors = [
        '000000', # 0
        'FFFFFF', # 1
        'FF0000', # 2
        '00FF00', # 3
        '0000FF', # 4
        'FFFF00', # 5
        'FF00FF', # 6
        '00FFFF', # 7
        '000000', # 8
        'FFFFFF', # 9
        'FF0000', # 10
        '00FF00', # 11
        '0000FF', # 12
        'FFFF00', # 13
        'FF00FF', # 14
        '00FFFF', # 15
        '800000', # 16
        '008000', # 17
        '000080', # 18
        '808000', # 19
        '800080', # 20
        '008080', # 21
        'C0C0C0', # 22
        '808080', # 23
        '9999FF', # 24
        '993366', # 25
        'FFFFCC', # 26
        'CCFFFF', # 27
        '660066', # 28
        'FF8080', # 29
        '0066CC', # 30
        'CCCCFF', # 31
        '000080', # 32
        'FF00FF', # 33
        'FFFF00', # 34
        '00FFFF', # 35
        '800080', # 36
        '800000', # 37
        '008080', # 38
        '0000FF', # 39
        '00CCFF', # 40
        'CCFFFF', # 41
        'CCFFCC', # 42
        'FFFF99', # 43
        '99CCFF', # 44
        'FF99CC', # 45
        'CC99FF', # 46
        'FFCC99', # 47
        '3366FF', # 48
        '33CCCC', # 49
        '99CC00', # 50
        'FFCC00', # 51
        'FF9900', # 52
        'FF6600', # 53
        '666699', # 54
        '969696', # 55
        '003366', # 56
        '339966', # 57
        '003300', # 58
        '333300', # 59
        '993300', # 60
        '993366', # 61
        '333399', # 62
        '333333'  # 63
      ]              

      attr_reader   :report
      attr_reader   :bindings

      def initialize (a_excel_filename, a_enable_cb_or_rb_edition=false, a_allow_sub_bands = true)
        super(a_excel_filename)
        @report_name                 = File.basename(a_excel_filename, '.xlsx')
        @current_band                = nil
        @first_row_in_band           = 0
        @band_type                   = nil
        @v_scale                     = 1
        @detail_cols_auto_height     = false
        @auto_float                  = false
        @auto_stretch                = false
        @band_split_type             = nil
        @basic_expressions           = false
        @allow_sub_bands             = a_allow_sub_bands
        @use_casper_bindings         = false
      end

      #
      # Convert XLS to VRXML.
      #
      # @param to Local file URI.
      #
      def convert(to: nil)

        read_all_tables()

        @report = JasperReport.new(@report_name)

        @binding = ::Xls::Vrxml::Binding.new(workbook: @workbook)
        @binding.load()

        @binding.map.each do | type, map |          

          # ... for all 'parameter/field/variable' ...
          map.each do | name, value |
            # ... ignore 'empty lines' ...
            next if '' == name
            # ... fetch binding ...
            binding = ::Xls::Vrxml::Binding.parse(type: type.to_s, value: value['Value'][:value] || "{\"__origin__\": \"#{__method__}\"}")
            case type
            when :parameters, :fields, :variables
              binding[:type] ||= 'String'
              binding[:java_class] ||= ::Xls::Vrxml::Binding.to_java_class(binding[:type])
            end
            # ... declare it ...
            case type
            when :parameters
              @report.parameters[name] = Parameter.new(name: name, binding: binding)
            when :fields
              @report.fields[name] = Field.new(name: name, binding: binding)
            when :variables
              @report.variables[name] = Variable.new(name: name, binding: binding)
            when :bands
              # nothing to do here
            when :named_cells
              @report.named_cells[name] = binding
            when :other
              case name
              # -
              when 'REPORT'
                binding.each do | k, v |
                  _attr = k.to_s.to_underscore
                  next if nil == v
                  if true == @report.respond_to?(_attr.to_sym)
                    @report.send("#{_attr}=", v)
                  else
                    Vrxml::Log.WHAT_IS(msg: "#{name} => #{k.to_s} = #{v}")
                  end
                  case k.to_s
                  when 'leftMargin'
                    @px_width = @report.page_width - @report.left_margin - @report.right_margin
                  when 'rightMargin'
                    @px_width = @report.page_width - @report.left_margin - @report.right_margin
                  else
                    # nothing to do
                  end
                end
              # -
              when 'GROUP'
                @report.group ||= Group.new
                binding.each do | k, v |
                  _attr = k.to_s.to_underscore
                  next if nil == v
                  if true == @report.respond_to?(_attr.to_sym)
                    @report.send("#{_attr}=", v)
                  else
                    case _attr
                    when 'expression'
                      @report.group.group_expression = v
                      declare_expression_entities(@report.group.group_expression)
                    else
                      Vrxml::Log.WHAT_IS(msg: "#{name} => #{k.to_s} = #{v}")
                    end
                  end
                end
              # -
              when 'OTHER'
                binding.each do | k, v |
                  _attr = k.to_s.to_underscore
                  next if nil == v 
                  if true == @report.respond_to?(_attr.to_sym)
                    @report.send("#{_attr}=", v)
                  else
                    case k
                    when :size
                      @report.paper_size = v  
                    when :vscale
                      @v_scale = v
                    when :query
                      @report.query_string = v
                    else
                      Vrxml::Log.WHAT_IS(msg: "#{name} => #{k.to_s} = #{v}")
                    end  
                  end
                  # update action(S)
                  case k
                  when :orientation
                    @report.update_page_size()
                  when :size
                    @report.update_page_size()
                  else
                    # nothing to do here
                  end
                end                
              else
                Vrxml::Log.WHAT_IS(msg: "#{name} => #{value['Value'][:value]}")
              end
            else
              ::Xls::Vrxml::Binding.halt(msg: " Don't know how to process '#{type}' binding!", file: __FILE__, line: __LINE__)
            end              
          end # map
        end # @binding.map

        # ... [B] consistency is not a requirement ... ...
        @layout_sheet_name = Xls::Vrxml::Object.guess_layout_sheet(workbook: @workbook)
        # ... [E] consistency is not a requirement ... ...

        # load named cells
        @name2ref, @ref2name = ::Xls::Vrxml::Binding.get_named_cells_map(sheet: ::Xls::Vrxml::Binding.get_sheet(named: @layout_sheet_name, at: @workbook), at: @workbook)
        

        @not_converted_expressions = {}
        @relationship              = 'lines'         


        generate_styles()

        @px_width = @report.page_width - @report.left_margin - @report.right_margin

        parse_sheets()

        #
        # Raise an error if we still have expressions to convert.
        #
        if false == @not_converted_expressions.empty?
          puts "---- ERROR ----".red
          @not_converted_expressions.each do |key, value|
            puts value[:jrxml]
            ap key
            puts value[:error]
          end
          raise "Unable to convert %d expression(s)" % [@not_converted_expressions.length]
        end                    
        #
        # TABLES
        #
        @@xmlns = "http://jasperreports.sourceforge.net/jasperreports"
        document = Nokogiri::XML(@report.to_xml)
        convert_to_tables(xml: document, relationship: @relationship)
        #
        # WRITE
        #
        xml = document.to_xml( :indent => 2, :encoding => Encoding::UTF_8.to_s, :save_with => Nokogiri::XML::Node::SaveOptions::AS_XML)
        return to.nil? ? xml : File.write(to, xml)
      end

      #
      # Convert to tables
      #
      # @param document
      # @param relationship
      #
      def convert_to_tables (xml:, relationship:)
        table 				= nil
        jasper_report = xml.xpath('x:jasperReport', 'x' => @@xmlns)
        detail_band   = xml.xpath('x:jasperReport/x:detail', 'x' => @@xmlns)
        column_header = xml.xpath('x:jasperReport/x:columnHeader', 'x' => @@xmlns)
        column_footer = xml.xpath('x:jasperReport/x:columnFooter', 'x' => @@xmlns)
        group				  = xml.xpath('x:jasperReport/x:group', 'x' => @@xmlns)      

        # log
        Vrxml::Log::LOG_IF(msg: "--- --- --- TABLES --- --- --- ", bits: Vrxml::Log::TABLES)
        
        # puts " - detail_band = #{detail_band.count}, column_header = #{column_header.count}, column_footer = #{column_footer.count}..."
        if 0 == detail_band.count && 0 == column_header.count && 0 == column_footer.count && 0 == group.count
          # log
          Vrxml::Log::LOG_IF(msg: "No tables...".white, bits: Vrxml::Log::TABLES)
        else
          # log
          Vrxml::Log::LOG_IF(msg: "Patching tables...".white, bits: Vrxml::Log::TABLES)
          # patch table
          jasper_report[0]['dataSourceType'] = 'legacy'

          if nil == table
            table = Nokogiri::XML::Node.new("table", xml)
            table['relationship'] = relationship
            if group.count > 0
              group[0].after(table)
            elsif column_header.count > 0
              column_header[0].after(table)
            elsif column_footer.count > 0
              column_footer[0].after(table)
            else # detail_band.count
              detail_band[0].after(table)
            end
          end
          # log
          Vrxml::Log::LOG_IF(msg: '  Re-parenting nodes...', bits: Vrxml::Log::TABLES)
          # 're-parenting' nodes
          [column_header, detail_band, column_footer, group].each do | a |
            a.each do | node |
              #
              node.parent = table
               # log
               Vrxml::Log::LOG_IF(msg: "    #{node.name}", bits: Vrxml::Log::TABLES)
              end
          end
        end # if

        # log
        Vrxml::Log::LOG_IF(msg: "--- --- --- --- --- --- --- ", bits: Vrxml::Log::TABLES)

      end # convert_to_tables

      private

      def generate_styles

        (0 .. @workbook.cell_xfs.size - 1).each do |style_index|
          style = xf_to_style(style_index)
          @report.add_style(name: style.name, value: style)
        end

      end

      def xf_to_style (a_style_index)

        # create a style
        style = Style.new('style_' + (a_style_index + 1).to_s)

        # grab cell format
        xf = @workbook.cell_xfs[a_style_index]

        # Format font
        if xf.apply_font == true
          xls_font = @workbook.fonts[xf.font_id]

          # on PDF we only have one font
          style.font_name = 'DejaVu Sans Condensed'
          # if xls_font.name.val == 'Arial'
          #   style.font_name = 'DejaVu Sans Condensed'
          # else
          #   style.font_name = xls_font.name.val
          # end

          unless xls_font.color.nil?
            style.forecolor = convert_color(xls_font.color)
          end

          style.font_size = xls_font.sz.val unless xls_font.sz.nil?
          style.is_bold   = true unless xls_font.b.nil?
          style.is_italic = true unless xls_font.i.nil?
        end

        # background
        if xf.apply_fill == true
          xls_fill = @workbook.fills[xf.fill_id]
          if xls_fill.pattern_fill.pattern_type == 'solid'
            style.backcolor = convert_color(xls_fill.pattern_fill.fg_color)
          end
        end

        # borders
        if xf.apply_border == true
          xls_border = @workbook.borders[xf.border_id]

          if xls_border.outline != nil

            if xls_border.outline.style != nil
              style.box ||= Box.new
              style.box.left_pen  = LeftPen.new
              style.box.top_pen   = TopPen.new
              style.box.right_pen = RightPen.new
              style.box.bottom    = BottomPen.new
              apply_border_style(style.box.left_pen  , xls_border.outline)
              apply_border_style(style.box.top_pen   , xls_border.outline)
              apply_border_style(style.box.right_pen , xls_border.outline)
              apply_border_style(style.box.bottom_pen, xls_border.outline)
            end

          else

            if xls_border.left != nil && xls_border.left.style != nil
              style.box ||= Box.new
              style.box.left_pen = LeftPen.new
              apply_border_style(style.box.left_pen, xls_border.left)
            end

            if xls_border.top != nil && xls_border.top.style != nil
              style.box ||= Box.new
              style.box.top_pen = TopPen.new
              apply_border_style(style.box.top_pen, xls_border.top)
            end

            if xls_border.right != nil && xls_border.right.style != nil
              style.box ||= Box.new
              style.box.right_pen = RightPen.new
              apply_border_style(style.box.right_pen, xls_border.right)
            end

            if xls_border.bottom != nil && xls_border.bottom.style != nil
              style.box ||= Box.new
              style.box.bottom_pen = BottomPen.new
              apply_border_style(style.box.bottom_pen, xls_border.bottom)
            end

          end
        end

        # Alignment
        if xf.apply_alignment

          #byebug if a_style_index == 111
          unless xf.alignment.nil?
            case xf.alignment.horizontal
            when 'left', nil
              style.h_text_align ='Left'
            when 'center'
              style.h_text_align ='Center'
            when 'right'
              style.h_text_align ='Right'
            end

            case xf.alignment.vertical
            when 'top'
              style.v_text_align ='Top'
            when 'center'
              style.v_text_align ='Middle'
            when 'bottom', nil
              style.v_text_align ='Bottom'
            end

            # rotation
            case xf.alignment.text_rotation
            when nil
              style.rotation = nil
            when 0
              style.rotation = 'None'
            when 90
              style.rotation = 'Left'
            when 180
              style.rotation = 'UpsideDown'
            when 270
              style.rotation = 'Right'
            end
          end
        end

        return style

      end

      def apply_border_style (a_pen, a_xls_border_style)
        case a_xls_border_style.style
        when 'thin'
          a_pen.line_width = 0.5
          a_pen.line_style = 'Solid'
        when 'medium'
          a_pen.line_width = 1.0
          a_pen.line_style = 'Solid'
        when 'dashed'
          a_pen.line_width = 1.0
          a_pen.line_style = 'Dotted'
        when 'dotted'
          a_pen.line_width = 0.5
          a_pen.line_style = 'Dotted'
        when 'thick'
          a_pen.line_width = 2.0
          a_pen.line_style = 'Solid'
        when 'double'
          a_pen.line_width = 0.5
          a_pen.line_style = 'Double'
        when 'hair'
          a_pen.line_width = 0.25
          a_pen.line_style = 'Solid'
        when 'mediumDashed'
          a_pen.line_width = 1.0
          a_pen.line_style = 'Dashed'
        when 'dashDot'
          a_pen.line_width = 0.5
          a_pen.line_style = 'Dashed'
        when 'mediumDashDot'
          a_pen.line_width = 1.0
          a_pen.line_style = 'Dashed'
        when 'dashDotDot'
          a_pen.line_width = 0.5
          a_pen.line_style = 'Dotted'
        when 'slantDashDot'
          a_pen.line_width = 0.5
          a_pen.line_style = 'Dotted'
        else
          a_pen.line_width = 1.0
          a_pen.line_style = 'Solid'
        end
        a_pen.line_color = convert_color(a_xls_border_style.color)
      end

      def convert_color (a_xls_color)
        if a_xls_color.indexed.nil?
          if a_xls_color.theme != nil
            cs = @workbook.theme.a_theme_elements.a_clr_scheme
            case a_xls_color.theme
            when 0
              return tint_theme_color(cs.a_lt1, a_xls_color.tint)
            when 1
              return tint_theme_color(cs.a_dk1, a_xls_color.tint)
            when 2
              return tint_theme_color(cs.a_lt2, a_xls_color.tint)
            when 3
              return tint_theme_color(cs.a_dk2, a_xls_color.tint)
            when 4
              return tint_theme_color(cs.a_accent1, a_xls_color.tint)
            when 5
              return tint_theme_color(cs.a_accent2, a_xls_color.tint)
            when 6
              return tint_theme_color(cs.a_accent3, a_xls_color.tint)
            when 7
              return tint_theme_color(cs.a_accent4, a_xls_color.tint)
            when 8
              return tint_theme_color(cs.a_accent5, a_xls_color.tint)
            when 9
              return tint_theme_color(cs.a_accent6, a_xls_color.tint)
            else
              return '#c0c0c0'
            end

          elsif a_xls_color.auto or a_xls_color.rgb.nil?
            return '#000000'
          else
            return '#' + a_xls_color.rgb[2..-1]
          end
        else
          return '#' + @@CT_IndexedColors[a_xls_color.indexed]
        end
      end

      def tint_theme_color (a_color, a_tint)
        color   = a_color.a_sys_clr.last_clr unless a_color.a_sys_clr.nil?
        color ||= a_color.a_srgb_clr.val
        r = color[0..1].to_i(16)
        g = color[2..3].to_i(16)
        b = color[4..5].to_i(16)
        unless a_tint.nil?
          if ( a_tint <  0 )
            a_tint = 1 + a_tint;
            r = r * a_tint
            g = g * a_tint
            b = b * a_tint
          else
            r = r + (a_tint * (255 - r))
            g = g + (a_tint * (255 - g))
            b = b + (a_tint * (255 - b))
          end
        end
        r = 255 if r > 255
        g = 255 if g > 255
        b = 255 if b > 255
        color = "#%02X%02X%02X" % [r, g, b]
        color
      end

      def parse_sheets
        @worksheet = nil
        @workbook.worksheets.each do |ws|
          if @layout_sheet_name != ws.sheet_name
            next
          end            
          @worksheet    = ws
          @raw_width    = 0
          @current_band = nil
          @band_type    = nil
          for col in (1 .. @worksheet.dimension.ref.col_range.end)
            @raw_width += get_column_width(@worksheet, col)
          end
          generate_bands()
        end
        raise "#{@layout_sheet_name} worksheet not found!" if nil == @worksheet
      end

      def generate_bands ()

        for row in @worksheet.dimension.ref.row_range
          next if @worksheet[row].nil?
          next if @worksheet[row][0].nil?
          row_tag = map_row_tag(@worksheet[row][0].value.to_s)
          next if row_tag.nil?

          if @band_type != row_tag
            adjust_band_height()
            process_row_mtag(row, row_tag)
            @first_row_in_band = row
          end
          unless @current_band.nil?
            generate_band_content(row)
          end
        end

        adjust_band_height()
      end

      def process_row_mtag (a_row, a_row_tag)
        if a_row_tag.nil? or a_row_tag.lines.size == 0
          process_row_tag(a_row, a_row_tag)
        else
          a_row_tag.lines.each do |tag|
            process_row_tag(a_row, tag)
          end
        end
      end

      def process_row_tag (a_row, a_row_tag)

        # grab cell info
        cell = { ref: RubyXL::Reference.ind2ref(a_row, 1).to_s }
        if @ref2name.include?(cell[:ref]) && @report.named_cells.include?(@ref2name[cell[:ref]])
          cell[:name] = @ref2name[cell[:ref]]
        end

        case a_row_tag
        when /CasperBinding:*/
          @use_casper_bindings = a_row_tag.split(':')[1].strip == 'true'
        when /BG\d*:/
          @report.background ||= Background.new
          @current_band = Band.new(tag: a_row_tag, cell: cell)
          @report.background.bands << @current_band
          @band_type = a_row_tag
        when /TL\d*:/
          @report.title ||= Title.new
          @current_band = Band.new(tag: a_row_tag, cell: cell)
          @report.title.bands << @current_band
          @band_type = a_row_tag
        when /PH\d*:/
          @report.page_header ||= PageHeader.new
          @current_band = Band.new(tag: a_row_tag, cell: cell)
          @report.page_header.bands << @current_band
          @band_type = a_row_tag
        when /CH\d*:/
          @report.column_header ||= ColumnHeader.new
          @current_band = Band.new(tag: a_row_tag, cell: cell)
          @report.column_header.bands << @current_band
          @band_type = a_row_tag
        when /DT\d*/
          @report.detail ||= Detail.new
          @current_band = Band.new(tag: a_row_tag, cell: cell)
          @report.detail.bands << @current_band
          @band_type = a_row_tag
        when /CF\d*:/
          @report.column_footer ||= ColumnFooter.new
          @current_band = Band.new(tag: a_row_tag, cell: cell)
          @report.column_footer.bands << @current_band
          @band_type = a_row_tag
        when /PF\d*:/
          @report.page_footer ||= PageFooter.new
          @current_band = Band.new(tag: a_row_tag, cell: cell)
          @report.page_footer.bands << @current_band
          @band_type = a_row_tag
        when /LPF\d*:/
          @report.last_page_footer ||= LastPageFooter.new
          @current_band = Band.new(tag: a_row_tag, cell: cell)
          @report.last_page_footer.bands << @current_band
          @band_type = a_row_tag
        when /SU\d*:/
          @report.summary ||= Summary.new
          @current_band = Band.new(tag: a_row_tag, cell: cell)
          @report.summary.bands << @current_band
          @band_type = a_row_tag
        when /ND\d*:/
          @report.no_data ||= NoData.new
          @current_band = Band.new(tag: a_row_tag, cell: cell)
          @report.no_data.bands << @current_band
          @band_type = a_row_tag
        when /GH\d*:/
          @report.group ||= Group.new
          @current_band = Band.new(tag: a_row_tag, cell: cell)
          @report.group.group_header.bands << @current_band
          @band_type = a_row_tag
        when /GF\d*:/
          @report.group ||= Group.new
          @current_band = Band.new(tag: a_row_tag, cell: cell)
          @report.group.group_footer.bands << @current_band
          @band_type = a_row_tag
        when /Id:.+/i
          @report.id = a_row_tag.split(':')[1].strip
        when /BasicExpressions:.+/i
          @widget_factory.basic_expressions = a_row_tag.split(':')[1].strip == 'true'
        when /Style:.+/i
          @current_band = nil
          @band_type    = nil
        else
          @current_band = nil
          @band_type    = nil
        end

        # band 'binding'
        if nil != @current_band && nil != @binding.map[:bands] && nil != @binding.map[:bands][@current_band.tag]
          obj = ::Xls::Vrxml::Binding.parse(type: :band, value:@binding.map[:bands][@current_band.tag]['Value'][:value])
          obj.each do | k, v |
            _attr = k.to_s.to_underscore
            if @current_band.respond_to?(_attr.to_sym)
              @current_band.send("#{_attr}=", v)
            else
              case k.to_s
              # TODO: 2.0 - EDITABLE
              # when 'lineParentIdField'
              #   @current_band.properties ||= Array.new
              #   @current_band.properties  << Property.new("epaper.casper.band.patch.op.add.attribute.name", value)
              # when 'dataRowTypeAttrName'
              #   @current_band.properties ||= Array.new
              #   @current_band.properties  << Property.new("epaper.casper.band.patch.op.add.attribute.data_row_type.name", value)
              when ''
              else
                ::Xls::Vrxml::Binding.halt(msg: "Don't know how to set '%s%s".yellow % [ "#{k.to_s}".red, "' attribute / property!".yellow ])
              end
            end
          end
        end # band 'binding'

      end

      def map_row_tag (a_row_tag)
        unless @allow_sub_bands
          match = a_row_tag.match(/\A(TL|SU|BG|PH|CH|DT|CF|PF|LPF|ND)\d*:\z/)
          if match != nil and match.size == 2
            return match[1] + ':'
          end
        end
        a_row_tag
      end

      def to_b (a_value)
        a_value.match(/(true|t|yes|y|1)$/i) != nil
      end

      def generate_band_content (a_row_idx)

        row = @worksheet[a_row_idx]

        max_cell_height = 0
        col_idx         = 1

        while col_idx < row.size do
                    
          col_span, row_span, cell_width, cell_height = measure_cell(a_row_idx, col_idx)

          if cell_width != nil

            if row[col_idx].nil? || row[col_idx].style_index.nil?
              col_idx += col_span
              next
            end

            field = create_field_legacy_mode(row[col_idx])
            field.report_element.x = x_for_column(col_idx)
            field.report_element.y = y_for_row(a_row_idx)
            field.report_element.width  = cell_width
            field.report_element.height = cell_height
            field.report_element.style  = 'style_' + (row[col_idx].style_index + 1).to_s


            if @current_band.stretch_type
              field.report_element.stretch_type = @current_band.stretch_type
            end

            if @current_band.auto_float and field.report_element.y != 0
              field.report_element.position_type = 'Float'
            end

            if @current_band.auto_stretch and field.respond_to?('is_stretch_with_overflow')
              field.is_stretch_with_overflow = true
            end

            # overide here with field by field directives
            process_field_comments(a_row_idx, col_idx, field)


            # If the field is from a horizontally merged cell we need to check the right side border
            if col_span > 1
              field.box ||= Box.new
              xf = @workbook.cell_xfs[row[col_idx + col_span - 1].style_index]
              if xf.apply_border
                xls_border = @workbook.borders[xf.border_id]

                if xls_border.right != nil && xls_border.right.style != nil
                  field.box ||= Box.new
                  field.box.right_pen = RightPen.new
                  apply_border_style(field.box.right_pen, xls_border.right)
                end
              end
            end

            # If the field is from a vertically merged cell we need to check the bottom side border
            if row_span > 1
              field.box ||= Box.new
              xf = @workbook.cell_xfs[@worksheet[a_row_idx + row_span - 1][col_idx].style_index]
              if xf.apply_border
                xls_border = @workbook.borders[xf.border_id]

                if xls_border.bottom != nil && xls_border.bottom.style != nil
                  field.box ||= Box.new
                  field.box.bottom_pen = BottomPen.new
                  apply_border_style(field.box.bottom_pen, xls_border.bottom)
                end
              end
            end
            if field_has_graphics(field)
              @current_band.children << field
              @report.style_set.add(field.report_element.style)
            end
          end
          col_idx += col_span
        end

      end

      def field_has_graphics (a_field)
        text_empty = false
        has_border = false
        opaque     = false

        if a_field.instance_of?(StaticText)
          if a_field.text.nil? || a_field.text.length == 0
            text_empty = true
          end
        end

        if a_field.instance_of?(TextField)
          if a_field.text_field_expression.nil? || a_field.text_field_expression.length == 0
            text_empty = true
          end
        end

        if a_field.box != nil
          if a_field.box.right_pen  != nil ||
              a_field.box.left_pen   != nil ||
              a_field.box.top_pen    != nil ||
              a_field.box.bottom_pen != nil
              has_border = true
          end
        end

        style = @report.styles[a_field.report_element.style]
        if style != nil
          if style.box != nil
            if style.box.right_pen  != nil ||
                style.box.left_pen   != nil ||
                style.box.top_pen    != nil ||
                style.box.bottom_pen != nil
              has_border = true
            end
          end
          if (style.mode != nil && style.mode == 'Opaque') || style.backcolor
            opaque = true
          end
        end

        return true if opaque

        return true if has_border

        return true unless text_empty

        return false
      end

      #
      # Obtain a property for a specific cell
      #
      # @param ref Cell reference.
      # @param property Symbol, property to read from binding table.
      #
      # @return Nil if not found.
      #
      def get_cell_binding_property(ref:, property:)
        if false == @report.named_cells.include?(ref)
          return nil
        end
        return @report.named_cells[ref], @report.named_cells[ref][property]
      end

      def create_field_legacy_mode (a_cell)
        f_id = nil
        rv  = nil
        binding = nil
        pattern = nil

        # for debug
        tracking = nil        
        if ::Xls::Vrxml::Log::DEBUG == ( ::Xls::Vrxml::Log::MASK & ::Xls::Vrxml::Log::DEBUG )
          tracking = Pathname.new(__FILE__).relative_path_from(Pathname.new(File.join(File.dirname(__FILE__), '../../..') )).to_s
        end
        # grab cell reference
        _ref = RubyXL::Reference.ind2ref(a_cell.row, a_cell.column)
        #
        cell = { ref: _ref.to_s }
        # basic text, no parameter(s)/field(s)/variable(s) or expression(s)
        if @ref2name.include?(_ref) && @report.named_cells.include?(@ref2name[_ref])
          cell[:name] = @ref2name[_ref]
        end
        # sanitize
        _exp = a_cell.value.to_s.strip
        if ( m = _exp.match(/\$SE\{(.*)\}/) )
          _exp = m[1]
        end
        # extract expression and related parameter(s)/field(s)/variable(s) ( if any )
        _exp, _ext = Vrxml::Expression.translate(expression: _exp, relationship: @relationship, nce: @not_converted_expressions)
        if _ext.count > 1
          # expression - contains parameter(s)/field(s)/variable(s)
          _ext.each do | e |
            case e[:type]
            when :parameter
              if false == @report.parameters.include?(e[:value])
                @report.add_parameter(id: e[:value], name: e[:value], java_class: nil)
              end
            when :field
              if false == @report.fields.include?(e[:value])
                @report.add_field(id: e[:value], name: e[:value], java_class: nil)
              end
            when :variable
              if false == @report.variables.include?(e[:value])
                @report.add_variable(id: e[:value], name: e[:value], java_class: nil)
              end
            else
              raise "#{e[:type]} - WTF?"
            end
          end
          # add text field element
          binding, patttern = get_cell_binding_property(ref: _ref, property: :pattern)
          if ::Xls::Vrxml::Log::DEBUG == ( ::Xls::Vrxml::Log::MASK & ::Xls::Vrxml::Log::DEBUG )
            tracking += ":#{__LINE__ + 2}"
          end
          rv = TextField.new(binding: binding, cell: cell, text_field_expression: _exp, pattern: pattern, tracking: tracking)
        elsif 1 == _ext.count
          # expression: single parameter/field/variable
          binding = nil
          pattern = nil
          case _ext[0][:type]
          when :parameter
            binding = @report.parameters[_ext[0][:value]] ? @report.parameters[_ext[0][:value]].binding : nil
            if nil != binding
              pattern = binding[:presentation]
            end
          when :field
            binding = @report.fields[_ext[0][:value]] ? @report.fields[_ext[0][:value]].binding : nil
            if nil != binding
              pattern = binding[:presentation]
            end
          when :variable
            binding = @report.variables[_ext[0][:value]] ? @report.variables[_ext[0][:value]].binding : nil
            if nil != binding
              pattern = binding[:presentation]
            end
          else
              raise "???"
          end
          # add text field element
          if ::Xls::Vrxml::Log::DEBUG == ( ::Xls::Vrxml::Log::MASK & ::Xls::Vrxml::Log::DEBUG )
            tracking += ":#{__LINE__ + 2}"
          end
          rv = TextField.new(binding: binding, cell: cell, text_field_expression: _exp, pattern: pattern, tracking: tracking)
        else        
          # basic text, no parameter(s)/field(s)/variable(s) or expression(s)
          if @ref2name.include?(_ref) && @report.named_cells.include?(@ref2name[_ref])
            binding, patttern = get_cell_binding_property(ref: @ref2name[_ref], property: :presentation)
          end
          # add text field element
          if ::Xls::Vrxml::Log::DEBUG == ( ::Xls::Vrxml::Log::MASK & ::Xls::Vrxml::Log::DEBUG )
            tracking += ":#{__LINE__ + ( nil != pattern ? 2 : 4 )}"
          end
          if nil != pattern
            rv = TextField.new(binding: binding, cell: cell, text_field_expression: _exp, pattern: pattern, tracking: tracking)
          else
            rv = StaticText.new(cell: cell, text: _exp, tracking: tracking)
          end
        end

        # TODO 2.0: implement
        if !f_id.nil? && rv.is_a?(TextField)
          ::Xls::Vrxml::Log.TODO(msg: "@ #{__method__}: implement @ #{__FILE__}:#{__LINE__} - java.util.Date")
          if @widget_factory.java_class(f_id) == 'java.util.Date'
            rv.text_field_expression = "DateFormat.parse(#{rv.text_field_expression},\"yyyy-MM-dd\")"
            rv.pattern_expression = "$P{i18n_date_format}"
            rv.report_element.properties << Property.new('epaper.casper.text.field.patch.pattern', 'yyyy-MM-dd') unless rv.report_element.properties.nil?
            parameter = Parameter.new(name: 'i18n_date_format', java_class: 'java.lang.String')
            parameter.default_value_expression = '"dd/MM/yyyy"'
            @report.parameters['i18n_date_format'] = parameter
          end
        end

        return rv
      end

      def declare_expression_entities (a_expression)

        all = Vrxml::Expression.extract(expression: a_expression) || []
        all.each do | element |
          f_id = element[:value]
          j_ks = nil # or 'java.lang.String'
          case element[:type]
          when :parameter
            @report.add_parameter(id: f_id, name: f_id, java_class: j_ks)
          when :field
            @report.add_field(id: f_id, name: f_id, java_class: j_ks)
          when :variable
            @report.add_variable(id: f_id, name: f_id, java_class: j_ks)
          else
            raise ArgumentError, "Don't know how to add '#{f_id}'!"
          end
        end
        nil
      end

      def process_field_comments (a_row, a_col, a_field)

        if @worksheet.comments != nil && @worksheet.comments.size > 0 && @worksheet.comments[0].comment_list != nil

          @worksheet.comments[0].comment_list.each do |comment|
            if comment.ref.col_range.begin == a_col && comment.ref.row_range.begin == a_row
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

                if tag == 'PE' or tag == 'printWhenExpression'
                  a_field.report_element.print_when_expression = transform_expression(expression: value) # to force declaration of paramters/fields/variables
                elsif tag == 'AF' or tag == 'autoFloat'
                  a_field.report_element.position_type = to_b(value) ? 'Float' : 'FixRelativeToTop'
                elsif tag == 'AS' or tag == 'autoStretch' and a_field.respond_to?(:is_stretch_with_overflow)
                  a_field.is_stretch_with_overflow = to_b(value)
                elsif tag == 'ST' or tag == 'stretchType'
                  a_field.report_element.stretch_type = value
                elsif tag == 'BN' or tag == 'blankIfNull' and a_field.respond_to?(:is_blank_when_null)
                  a_field.is_blank_when_null = to_b(value)
                elsif tag == 'PT' or tag == 'pattern' and a_field.respond_to?(:pattern)
                  a_field.pattern = value
                elsif tag == 'ET' or tag == 'evaluationTime' and a_field.respond_to?(:evaluation_time)
                  a_field.evaluation_time = value.capitalize
                elsif tag == 'DE' or tag == 'disabledExpression'
                  _exp = transform_expression(expression: value)
                  if a_field.respond_to? :disabled_conditional
                    a_field.disabled_conditional(_exp)
                  else
                    a_field.report_element.properties ||= Array.new
                    a_field.report_element.properties << PropertyExpression.new('epaper.casper.text.field.disabled.if', _exp)
                  end
                elsif tag == 'SE' or tag == 'styleExpression'
                  _exp = transform_expression(expression: value) # to force declaration of parameters/fields/variables
                  if a_field.respond_to? :style_expression
                    a_field.style_expression(_exp)
                  else
                    a_field.report_element.properties ||= Array.new
                    a_field.report_element.properties << PropertyExpression.new('epaper.casper.style.condition', _exp)
                  end
                elsif tag == 'RIC' or tag == 'reloadIfChanged'
                  if a_field.respond_to? :reload_if_changed
                    a_field.reload_if_changed(value)
                  else
                    a_field.report_element.properties ||= Array.new
                    a_field.report_element.properties << Property.new('epaper.casper.text.field.reload.if_changed', value)
                  end
                elsif tag == 'EE' or tag == 'editableExpression'
                  _exp = transform_expression(expression: value) # to force declaration of parameters/fields/variables
                  if a_field.respond_to? :enabled_conditional
                    a_field.enabled_conditional(_exp)
                  else
                    a_field.report_element.properties ||= Array.new
                    a_field.report_element.properties << PropertyExpression.new('epaper.casper.text.field.editable.if', _exp)
                  end
                end
              end

            end
          end
        end
      end

      def transform_expression(expression:)
        _exp, _ext = Vrxml::Expression.translate(expression: v2, relationship: @relationship, nce: @not_converted_expressions)
        # add all parameters/fields/variables
        if _ext.count > 0
          _ext.each do | element |
            f_id = element[:value]
            j_ks = nil # or 'java.lang.String'
            case element[:type]
            when :parameter
              @report.add_parameter(id: f_id, name: f_id, java_class: j_ks)
            when :field
              @report.add_field(id: f_id, name: f_id, java_class: j_ks)
            when :variable
              @report.add_variable(id: f_id, name: f_id, java_class: j_ks)
            else
              raise "WTF ???"
            end
          end
        end
        # done
        return _exp
      end

      def get_column_width (a_worksheet, a_index)
        width   = a_worksheet.get_column_width_raw(a_index)
        width ||= RubyXL::ColumnRange::DEFAULT_WIDTH
        return width
      end


      def x_for_column (a_col_idx)

        width = 0
        for idx in (1 .. a_col_idx - 1) do
          width += get_column_width(@worksheet, idx)
        end
        return scale_x(width).round

      end

      def y_for_row (a_row_idx)
        height = 0
        for idx in (@first_row_in_band .. a_row_idx - 1) do
          height += @worksheet.get_row_height(idx)
        end
        return scale_y(height).round
      end

      def adjust_band_height ()

        return if @current_band.nil?

        height = 0
        for row in @worksheet.dimension.ref.row_range
          unless @worksheet[row].nil? or @worksheet[row][0].nil? or @worksheet[row][0].value.nil? or map_row_tag(@worksheet[row][0].value) != @band_type
            height += y_for_row(row + 1) - y_for_row(row)
          end
        end

        @current_band.height = height
      end

      def measure_cell (a_row_idx, a_col_idx)

        @worksheet.merged_cells.each do |merged_cell|

          col_span = merged_cell.ref.col_range.size
          row_span = merged_cell.ref.row_range.size

          if a_row_idx == merged_cell.ref.row_range.begin && a_col_idx == merged_cell.ref.col_range.begin

            cell_height = y_for_row(merged_cell.ref.row_range.end + 1) -  y_for_row(merged_cell.ref.row_range.begin)
            cell_width  = x_for_column(merged_cell.ref.col_range.end + 1) - x_for_column(merged_cell.ref.col_range.begin)

            return col_span, row_span, cell_width, cell_height

          elsif merged_cell.ref.row_range.include?(a_row_idx) and merged_cell.ref.col_range.include?(a_col_idx)

            # The cell is overlaped by a merged cell
            return col_span, row_span, nil, nil

          end
        end

        cell_height = y_for_row(a_row_idx + 1) -  y_for_row(a_row_idx)
        cell_width  = x_for_column(a_col_idx + 1) - x_for_column(a_col_idx)
        return 1, 1, cell_width, cell_height

      end

      def scale_x (a_width)
        return (a_width * @px_width / @raw_width)
      end

      def scale_y (a_height)
        return (a_height * @v_scale)
      end


    end # class Converter

  end # of module 'Vrxml'
end # of module 'Xls'
