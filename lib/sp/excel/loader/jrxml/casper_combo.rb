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

        class CasperCombo < CasperTextField

          def initialize (a_generator, a_expression)
            super(a_generator, a_expression)

            if @binding.cc_field_name[0] != '['
              field_name = @binding.cc_field_name
              fields     = [ @binding.cc_field_id, @binding.cc_field_name ]
            else
              field_name = @binding.cc_field_name[1..-2]
              fields     = @binding.cc_field_name[1..-2].split(',').each { |e| e.strip! }
            end

            if @binding.respond_to?(:html)
              html = @binding.html
            else
              html = "<div class=\"normal\"><div class=\"left\">[[#{fields[0]}]]</div><div class=\"main\">[[#{fields[1]}]]</div></div>"
            end

            if @binding.respond_to?(:cc_field_patch) and @binding.cc_field_patch != ''
              patch_name = @binding.cc_field_patch
            else
              patch_name = @binding.id[3..-2]
            end

            @casper_binding[:editable] = {
                is: @binding.editable,
                patch: {
                  field: {
                    type: @binding.java_class,
                    name: patch_name
                  }
                }
              }

            @casper_binding[:attachment] = {
                type: 'dropDownList',
                version: 2,
                controller: 'client',
                route: @binding.uri.gsub('"', '""'),
                display: fields,
                html: html
              }

            if @binding.respond_to?(:allow_clear)
              unless @binding.allow_clear.nil? or not @binding.allow_clear
                @casper_binding[:attachment][:allowClear] = @binding.allow_clear
              end
            end

          end

        end
      end
    end
  end
end