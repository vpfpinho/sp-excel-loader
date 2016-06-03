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

module Sp
  module Excel
    module Loader
      module Jrxml

        class TextField < StaticText

          attr_accessor :text_field_expression
          attr_accessor :is_stretch_with_overflow
          attr_accessor :is_blank_when_null
          attr_accessor :pattern

          def initialize(a_properties, a_pattern = nil, a_pattern_expression = nil)
            super()
            @text_field_expression     = nil
            @is_blank_when_null        = nil
            @is_stretch_with_overflow  = false
            @pattern                   = a_pattern
            @pattern_expression        = a_pattern_expression
            @report_element.properties = a_properties
          end

          def attributes
            rv = Hash.new
            rv['isStretchWithOverflow'] = true if @is_stretch_with_overflow
            rv['pattern']               = @pattern unless @pattern.nil?
            rv['isBlankWhenNull']       = @is_blank_when_null unless @is_blank_when_null.nil?
            return rv
          end

          def to_xml (a_node)
            Nokogiri::XML::Builder.with(a_node) do |xml|
              xml.textField(attributes)
            end
            @report_element.to_xml(a_node.children.last)
            @box.to_xml(a_node.children.last) unless @box.nil?
            if nil != @text_field_expression && @text_field_expression.length > 0
              Nokogiri::XML::Builder.with(a_node.children.last) do |xml|
                xml.textFieldExpression {
                  xml.cdata(@text_field_expression)
                }
              end
            end
            if nil != @pattern_expression && @pattern_expression.length > 0
              Nokogiri::XML::Builder.with(a_node.children.last) do |xml|
                xml.patternExpression {
                  xml.cdata(@pattern_expression)
                }
              end
            end
          end

        end

      end
    end
  end
end
