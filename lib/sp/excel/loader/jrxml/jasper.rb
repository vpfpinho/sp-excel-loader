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
          # Report class instance data
          #
          attr_accessor :parameters
          attr_accessor :styles
          attr_accessor :style_set
          attr_accessor :fields
          attr_accessor :variables
          attr_accessor :band_containers
          attr_accessor :builder
          attr_accessor :group
          attr_accessor :detail
          attr_accessor :query_string
          attr_accessor :page_width
          attr_accessor :page_height
          attr_accessor :no_data_section
          attr_accessor :column_width
          attr_accessor :left_margin
          attr_accessor :right_margin
          attr_accessor :top_margin
          attr_accessor :bottom_margin
          attr_accessor :report_name
          attr_accessor :orientation
          attr_accessor :paper_size
          attr_accessor :properties

          def initialize (a_name)

            # init data set
            @group           = nil
            @detail          = nil
            @query_string    = nil
            @parameters      = Hash.new
            @fields          = Hash.new
            @variables       = Hash.new
            @styles          = Hash.new
            @style_set       = Set.new
            @band_containers = Array.new

            # defaults for jasper report attributes
            @orientation     = 'Landscape'
            @paper_size      = 'A4'
            @page_width      = 595
            @page_height     = 842
            @no_data_section = 'NoPages'
            @column_width    = 522
            @left_margin     = 36
            @right_margin    = 37
            @top_margin      = 30
            @bottom_margin   = 30
            @report_name     = a_name
            @is_summary_with_page_header_and_footer = true;
            @is_float_column_footer                 = true;
            @generator_version = Sp::Excel::Loader::VERSION
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
                               'isSummaryWithPageHeaderAndFooter' => @is_summary_with_page_header_and_footer,
                               'isFloatColumnFooter'              => @is_float_column_footer) {
                xml.comment('created with core-excel-loader ' + @generator_version)
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
              variable.to_xml(@builder.doc.children[0])
            end

            @group.to_xml(@builder.doc.children[0]) unless @group.nil?

            summary_bands    = @band_containers.reject { |e| "SU" != e.band_type }
            all_other_bands  = @band_containers.reject { |e| "SU" == e.band_type }
            @band_containers = all_other_bands + summary_bands

            @band_containers.each do |band_container|
              band_container.to_xml(@builder.doc.children[0])
            end

            @builder.to_xml(indent:2)
          end

        end

      end
    end
  end
end
