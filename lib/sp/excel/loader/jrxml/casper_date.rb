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

          def initialize (a_generator, a_expression)
            super(a_generator, a_expression)
            a_generator.declare_expression_entities(a_expression)

            @casper_binding[:editable] = {
                patch: {
                  field: {
                    pattern: 'yyyy-MM-dd'
                  }
                }                        
              }            
              
            @casper_binding[:attachment] = {
                type: 'datePicker',
                version: 2
              }

            @text_field_expression = "DateFormat.parse(#{a_expression},\"yyyy-MM-dd\")"
            @pattern_expression = '$P{i18n_date_format}'
          end

        end

      end
    end
  end
end