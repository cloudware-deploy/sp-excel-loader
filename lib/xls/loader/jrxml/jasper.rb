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

module Xls
  module Loader
    module Jrxml

      class JasperReport

        #
        # Attributes that can be configured using row tags
        #
        attr_accessor :column_width
        attr_accessor :left_margin
        attr_accessor :right_margin
        attr_accessor :top_margin
        attr_accessor :bottom_margin
        attr_accessor :report_name
        attr_accessor :is_title_new_page
        attr_accessor :data_source_type

        #
        # Report class instance data
        #
        attr_reader :parameters
        attr_reader :fields
        attr_reader :variables
        attr_reader :named_cells

        attr_reader :styles

        attr_accessor :style_set
        attr_accessor :builder
        attr_accessor :group
        attr_accessor :query_string
        attr_accessor :page_width
        attr_accessor :page_height
        attr_accessor :no_data_section
        attr_accessor :orientation
        attr_accessor :paper_size
        attr_accessor :properties
        attr_reader   :extension

        # band containers
        attr_accessor :detail
        attr_accessor :title
        attr_accessor :background
        attr_accessor :page_header
        attr_accessor :column_header
        attr_accessor :column_footer
        attr_accessor :page_footer
        attr_accessor :last_page_footer
        attr_accessor :summary
        attr_accessor :no_data

        def initialize (a_name)

          # init data set
          @group            = nil
          @detail           = nil
          @title            = nil
          @background       = nil
          @page_header      = nil
          @column_header    = nil
          @column_footer    = nil
          @page_footer      = nil
          @last_page_footer = nil
          @summary          = nil
          @no_data          = nil

          @query_string    = 'lines'
          @parameters      = Hash.new
          @fields          = Hash.new
          @variables       = Hash.new
          @styles          = Hash.new
          @style_set       = Set.new

          @named_cells    = Hash.new

          # defaults for jasper report attributes
          @orientation       = 'Portrait'
          @paper_size        = 'A4'
          @page_width        = 595
          @page_height       = 842
          @no_data_section   = 'NoPages'
          @column_width      = 522
          @left_margin       = 36
          @right_margin      = 37
          @top_margin        = 30
          @bottom_margin     = 30
          @report_name       = a_name
          @is_title_new_page = false
          @is_summary_with_page_header_and_footer = true;
          @is_float_column_footer                 = true;
          @generator_version = Xls::Loader::VERSION.strip
          @fields["$['lines'][index]['data_row_type']"] = Field.new(name: "$['lines'][index]['data_row_type']", 
            binding: JSON.parse({type: 'integer', java_class: ::Xls::Vrxml::Binding.to_java_class('integer')}.to_json, symbolize_names: true)
          )
          Variable.known_variables.each do | definition, |
            @variables["$.$$VARIABLES[index]['#{definition[:name]}']"] = Variable.new(name: "$.$$VARIABLES[index]['#{definition[:name]}']", 
              binding: JSON.parse(definition.to_json, symbolize_names: true)
            )
          end
          @extension = nil # TODO 2.0: REMOVE ? ReportExtension.new(@report_name)
          #
          @data_source_type = 'legacy'
        end

        def update_page_size
          case @paper_size
          when 'A4'
            if @orientation == 'Landscape'
              @page_width  = 842
              @page_height = 595
            else
              @page_width  = 595
              @page_height = 842
            end
          else
            @page_width  = 595
            @page_height = 842
          end
        end

        def update_extension_style (a_name, a_cell)
          @extension.styles.delete_if {|style| style.name == a_name}
          style              = @styles.delete("style_#{a_cell.style_index+1}")
          style.name         = a_name
          style.v_text_align = nil
          style.h_text_align = nil
          @styles[a_name] = style
          @style_set.add(a_name)
        end

        def to_xml
          @builder = Nokogiri::XML::Builder.new(:encoding => 'UTF-8') do |xml|
            xml.jasperReport('xmlns'              => 'http://jasperreports.sourceforge.net/jasperreports',
                              'xmlns:xsi'          => 'http://www.w3.org/2001/XMLSchema-instance',
                              'xsi:schemaLocation' => 'http://jasperreports.sourceforge.net/jasperreports http://jasperreports.sourceforge.net/xsd/jasperreport.xsd',
                              'name'               => @report_name,
                              'pageWidth'          => @page_width,
                              'pageHeight'         => @page_height,
                              'whenNoDataType'     => @no_data_section,
                              'columnWidth'        => @column_width,
                              'leftMargin'         => @left_margin,
                              'rightMargin'        => @right_margin,
                              'topMargin'          => @top_margin,
                              'bottomMargin'       => @bottom_margin,
                              'isTitleNewPage'     => @is_title_new_page,
                              'isSummaryWithPageHeaderAndFooter' => @is_summary_with_page_header_and_footer,
                              'isFloatColumnFooter'              => @is_float_column_footer,
                              'dataSourceType' => @data_source_type
            ) {
              xml.comment('created with xls2vexml ' + @generator_version)
            }
          end

          # TODO 2.0: remove ?
          # if not @extension.nil?

          #   if not @extension.properties.nil?
          #     @extension.properties.each do |property|
          #       property.to_xml(@builder.doc.children[0])
          #     end
          #   end

          #   if not @extension.styles.nil?
          #     @extension.styles.each do |style|
          #       style.to_xml(@builder.doc.children[0])
          #     end
          #   end
          # end

          #
          # WRITE STYLES
          #
          Nokogiri::XML::Builder.with(@builder.doc.children[0]) do |xml|
            xml.comment(" STYLES ")
          end          
          @styles.each do |name, style|
            if @style_set.include? name
              style.to_xml(@builder.doc.children[0])
            end
          end

          #
          # WRITE PARAMETERS
          #
          Nokogiri::XML::Builder.with(@builder.doc.children[0]) do |xml|
            xml.comment(" PARAMETERS ")
          end          
          @parameters.each do |name, parameter|
            parameter.to_xml(@builder.doc.children[0])
          end

          # TODO 2.0: remove?
          # unless @query_string.nil?
          #   Nokogiri::XML::Builder.with(@builder.doc.children[0]) do |xml|
          #     xml.queryString {
          #       xml.cdata(@query_string)
          #     }
          #   end
          # end

          #
          # WRITE FIELDS
          #
          Nokogiri::XML::Builder.with(@builder.doc.children[0]) do |xml|
            xml.comment(" FIELDS ")
          end          
          @fields.each do |name, field|
            field.to_xml(@builder.doc.children[0])
          end

          #
          # WRITE VARIABLES
          #
          Nokogiri::XML::Builder.with(@builder.doc.children[0]) do |xml|
            xml.comment(" VARIABLES ")
          end          
          @variables.each do |name, variable|
            next if ['PAGE_NUMBER', 'MASTER_CURRENT_PAGE', 'MASTER_TOTAL_PAGES', 
                      'COLUMN_NUMBER', 'REPORT_COUNT', 'PAGE_COUNT', 'COLUMN_COUNT'].include? name
            variable.to_xml(@builder.doc.children[0])
          end

          #
          # WRITE LAYOUT
          #
          Nokogiri::XML::Builder.with(@builder.doc.children[0]) do |xml|
            xml.comment(" LAYOUT ")
          end

          # group
          @group.to_xml(@builder.doc.children[0]) unless @group.nil?

          # other
          @background.to_xml(@builder.doc.children[0])       unless @background.nil?
          @title.to_xml(@builder.doc.children[0])            unless @title.nil?
          @page_header.to_xml(@builder.doc.children[0])      unless @page_header.nil?
          @column_header.to_xml(@builder.doc.children[0])    unless @column_header.nil?
          @detail.to_xml(@builder.doc.children[0])           unless @detail.nil?
          @column_footer.to_xml(@builder.doc.children[0])    unless @column_footer.nil?
          @page_footer.to_xml(@builder.doc.children[0])      unless @page_footer.nil?
          @last_page_footer.to_xml(@builder.doc.children[0]) unless @last_page_footer.nil?
          @summary.to_xml(@builder.doc.children[0])          unless @summary.nil?
          @no_data.to_xml(@builder.doc.children[0])          unless @no_data.nil?

          #
          # finalize
          #
          @builder.to_xml(indent:2)
        end

        #
        # Add a parameter.
        #
        # @param id     ID
        # @param name   Common name.
        # @param caller For debug purpose only.
        #
        def add_parameter(id:, name:, java_class:, caller: caller_locations(1,1)[0].base_label)
          if @parameters.has_key?(name)
            return
          end
          
          parameter = Parameter.new(name: name, java_class: java_class)

          Vrxml::Log.TODO(msg: "#{__method__} called from #{caller} - NEEDS REVIEW")

          # if @bindings.has_key? id
          #   binding = @bindings[id]
          #   if binding.respond_to? 'default' and binding.default != nil and binding.default.strip != ''
          #     if binding.java_class == 'java.lang.String'
          #       parameter.default_value_expression = "\"#{binding.default.strip}\""
          #     else
          #       parameter.default_value_expression = binding.default.strip
          #     end
          #   end
          # end
          @parameters[name] = parameter
        end

        #
        # Add a field.
        #
        # @param id     ID
        # @param name   Common name.
        # @param caller For debug purpose only.
        #
        def add_field(id:, name:, java_class:, caller: caller_locations(1,1)[0].base_label)
          if @fields.has_key?(name)
            return
          end
          
          Vrxml::Log.TODO(msg: "#{__method__} called from #{caller} - NEEDS REVIEW")

          @fields[name] = Field.new(name: name, java_class: java_class)
        end

        #
        # Add a variable.
        #
        # @param id     ID
        # @param name   Common name.
        # @param caller For debug purpose only.
        #
        def add_variable(id:, name:, java_class:, caller: caller_locations(1,1)[0].base_label)
          if "PAGE_NUMBER" == name || @report.variables.has_key?(name)
            return
          end

          Vrxml::Log.TODO(msg: "#{__method__} called from #{caller} - NEEDS REVIEW")

          @report.variables[name] = Variable.new(name: name, java_class: java_class)
        end

        #
        # Add a style.
        #
        # @param name  Unique name ( same as ID ).
        # @param value Properties.
        #
        def add_style(name:, value:)
          @styles[name] = value
        end

      end # class 'JasperReport'

    end
  end
end
