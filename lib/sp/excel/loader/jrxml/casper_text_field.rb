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

        class CasperTextField < TextField

          attr_reader :casper_binding

          def initialize (a_bindings, a_generator, a_expression)
            super(Array.new, nil, nil)
            @value_types = Array.new
            @casper_binding = {}
          end

          def convert_type a_value
            case a_value.to_s
            when /\A".+"\z/
              rv = a_value[1..-1]
              @value_types << 'java.lang.String'
            when 'true'
              rv = true
              @value_types << 'java.lang.Boolean'
            when 'false'
              rv = false
              @value_types << 'java.lang.Boolean'
            when /([+-])?\d+/
              rv = Integer(a_value) rescue nil
              @value_types << 'java.lang.Integer'
            else
              rv = Float(a_value) rescue nil
              @value_types << 'java.lang.Double'
            end
            if rv.nil?
              raise "Unable to convert value #{a_value} to json"
            end
            rv
          end


          def disabled_expression (value)
            #a_field.report_element.properties << PropertyExpression.new('epaper.casper.text.field.disabled.if', value)
          end

          def style_expression (value)
            #a_field.report_element.properties << PropertyExpression.new('epaper.casper.style.condition', value)
          end
          
          def reload_if_changed (value)
            #a_field.report_element.properties << Property.new('epaper.casper.text.field.reload.if_changed', value)
          end
            
          def editable_expression (value)
            # a_field.report_element.properties << PropertyExpression.new('epaper.casper.text.field.editable.if', value)
          end

        end

      end
    end
  end
end