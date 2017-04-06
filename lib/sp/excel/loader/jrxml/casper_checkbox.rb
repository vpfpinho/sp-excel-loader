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

        class CasperCheckbox < CasperTextField

          def validation_regexp 
            /\A\$CB{(\$[PFV]{.+}),(.+),(.+)}\z/
          end

          def attachment
            'checkbox'
          end

          #
          # check box: $CB{<field_name>,<unchecked>,<checked>}
          #
          def initialize (a_generator, a_expression)
            super(a_generator, a_expression)
            
            # validade expression and extract components
            values = validation_regexp.match a_expression.delete(' ')
            if values.nil? or values.size != 4 
              raise "Invalid checkbox expression: '#{a_expression}'"
            else
              field_expr = values[1]
              on_value   = convert_type(values[2])
              off_value  = convert_type(values[3])
            end

            a_generator.declare_expression_entities(a_expression)

            # get or guess the expression type
            binding = a_generator.bindings[field_expr]
            if not binding.nil?
              type = binding.java_class 
              if type != @value_types[0] 
                raise "Checked value '#{on_value}' type #{@value_types[0]} does not match the binding type #{type} (#{a_expression})"
              end
              if type != @value_types[1]
                raise "Checked value '#{off_value}' type #{@value_types[1]} does not match the binding type #{type} (#{a_expression})"
              end
            else
              if @value_types[0] != @value_types[1]
                raise "Checked value '#{on_value}' type #{@value_types[0]} is different from unchecked '#{off_value}' type #{@value_types[1]} (#{a_expression})"
              end
              type = @value_types[0]
            end

            @casper_binding = {
              editable: {
                patch: {
                  field: {
                    name: field_expr[3..-2]  
                  }
                }
              },
              attachment: {
                type: attachment,
                version: 2,
                value: {
                  type: type,
                  on: on_value,
                  off: off_value
                }
              }
            }

            @text_field_expression = "#{field_expr} == #{on_value} ? \"X\" : \"\""

          end

        end

      end
    end
  end
end