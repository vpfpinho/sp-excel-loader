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

        class CasperDate < CasperTextField

          def initialize (a_binding, a_generator, expression)
            super(Array.new, a_binding.presentation.format, nil)

            a_generator.declare_expression_entities(expression)

            binding = {
                        editable: {
                          patch: {
                            field: {
                              pattern: 'yyyy-MM-dd'
                            }
                          }                        
                        },             
                        attachment: {
                          type: 'datePicker',
                          version: 2,
                        }
                      }

            unless a_binding.tooltip.nil? or a_binding.tooltip.empty?
              binding[:hint][:expression] = a_binding.tooltip
              a_generator.declare_expression_entities(a_binding.tooltip)
            end

            @text_field_expression = "DateFormat.parse(#{expression},\"yyyy-MM-dd\")"
            @pattern_expression = '$P{i18n_date_format}'

            #if !f_id.nil? && rv.is_a?(TextField)
            #  if @widget_factory.java_class(f_id) == 'java.util.Date'
            #    rv.text_field_expression = "DateFormat.parse(#{rv.text_field_expression},\"yyyy-MM-dd\")"
            #    rv.pattern_expression = "$P{i18n_date_format}"
            #    rv.report_element.properties << Property.new('epaper.casper.text.field.patch.pattern', 'yyyy-MM-dd') unless rv.report_element.properties.nil?
            #    parameter = Parameter.new('i18n_date_format', 'java.lang.String')
            #    parameter.default_value_expression = '"dd/MM/yyyy"'
            #    @report.parameters['i18n_date_format'] = parameter
            #  end
            #end
            ap binding
            @report_element.properties << Property.new('casper.binding', binding.to_json)

          end

        end

      end
    end
  end
end