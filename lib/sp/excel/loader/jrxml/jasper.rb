# encoding: utf-8
#
# Copyright (c) 2011-2016 Servicepartner LDA. All rights reserved.
#
# This file is part of sp-excel-loader.
#
# sp-excel-loader is free software: you can redistribute it and/or modify
# it under the terms of the GNU Affero General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# sp-excel-loader is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU Affero General Public License
# along with sp-excel-loader.  If not, see <http://www.gnu.org/licenses/>.
#

require 'set'

module Sp
  module Excel
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

          #
          # Report class instance data
          #
          attr_accessor :parameters
          attr_accessor :styles
          attr_accessor :style_set
          attr_accessor :fields
          attr_accessor :variables
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
          attr_accessor :title
          attr_accessor :background
          attr_accessor :page_header
          attr_accessor :page_footer
          attr_accessor :last_page_footer
          attr_accessor :summary
          attr_accessor :no_data
          attr_accessor :body

          def initialize (a_name)

            # init data set
            #@group            = nil
            @title            = nil
            @background       = nil
            @page_header      = nil
            @page_footer      = nil
            @last_page_footer = nil
            @summary          = nil
            @no_data          = nil
            @body             = []

            @query_string    = 'lines'
            @parameters      = Hash.new
            @fields          = Hash.new
            @variables       = Hash.new
            @styles          = Hash.new
            @style_set       = Set.new

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
            @generator_version = Sp::Excel::Loader::VERSION.strip
            @fields['data_row_type'] = Field.new('data_row_type')
            @variables['ON_LAST_PAGE'] = Variable.new('ON_LAST_PAGE', 'java.lang.Boolean')

            @extension = ReportExtension.new(@report_name)

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
                               'isFloatColumnFooter'              => @is_float_column_footer) {
                xml.comment('created with sp-excel-loader ' + @generator_version)
              }
            end

            if not @extension.nil?

              if not @extension.properties.nil?
                @extension.properties.each do |property|
                  property.to_xml(@builder.doc.children[0])
                end
              end

              if not @extension.styles.nil?
                @extension.styles.each do |style|
                  style.to_xml(@builder.doc.children[0])
                end
              end
            end

            @styles.each do |name, style|
              if @style_set.include? name
                style.to_xml(@builder.doc.children[0])
              end
            end

            @parameters.each do |name, parameter|
              parameter.to_xml(@builder.doc.children[0])
            end

            unless @query_string.nil?
              Nokogiri::XML::Builder.with(@builder.doc.children[0]) do |xml|
                xml.queryString {
                  xml.cdata(@query_string)
                }
              end
            end

            @fields.each do |name, field|
              field.to_xml(@builder.doc.children[0])
            end

            @variables.each do |name, variable|
              next if ['PAGE_NUMBER', 'MASTER_CURRENT_PAGE', 'MASTER_TOTAL_PAGES',
                       'COLUMN_NUMBER', 'REPORT_COUNT', 'PAGE_COUNT', 'COLUMN_COUNT'].include? name
              variable.to_xml(@builder.doc.children[0])
            end

            #@group.to_xml(@builder.doc.children[0]) unless @group.nil?

            @background.to_xml(@builder.doc.children[0])       unless @background.nil?
            @title.to_xml(@builder.doc.children[0])            unless @title.nil?
            @page_header.to_xml(@builder.doc.children[0])      unless @page_header.nil?
            @body.each do |part|
              part.to_xml(@builder.doc.children[0])
            end
            @page_footer.to_xml(@builder.doc.children[0])      unless @page_footer.nil?
            @last_page_footer.to_xml(@builder.doc.children[0]) unless @last_page_footer.nil?
            @summary.to_xml(@builder.doc.children[0])          unless @summary.nil?
            @no_data.to_xml(@builder.doc.children[0])          unless @no_data.nil?

            @builder.to_xml(indent:2)
          end

        end

      end
    end
  end
end
