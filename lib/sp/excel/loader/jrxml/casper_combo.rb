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

            if a_binding.cc_field_name[0] != '['
              field_name = a_binding.cc_field_name
              fields     = [ a_binding.cc_field_id, a_binding.cc_field_name ]
            else
              field_name = a_binding.cc_field_name[1..-2]
              fields     = a_binding.cc_field_name[1..-2].split(',').each { |e| e.strip! }
            end

            if a_binding.respond_to?(:html)
              html = a_binding.html
            else
              html = "<div class=\"normal\"><div class=\"left\">[[#{fields[0]}]]</div><div class=\"main\">[[#{fields[1]}]]</div></div>"
            end

            if a_binding.respond_to?(:cc_field_patch) and a_binding.cc_field_patch != ''
              patch_name = a_binding.cc_field_patch
            else
              patch_name = a_binding.id[3..-2]
            end

            @casper_binding = {
                        editable: {
                          patch: {
                            field: {
                              type: a_binding.java_class,
                              name: patch_name
                            }
                          }
                        },
                        attachment: {
                          type: 'dropDownList',
                          version: 2,
                          controller: 'client',
                          route: a_binding.uri.gsub('"', '""'),
                          display: fields,
                          html: html
                        }
                      }

            if a_binding.respond_to?(:allow_clear) 
              unless a_binding.allow_clear.nil? or !a_binding.allow_clear
                @casper_binding[:attachment][:allowClear] = a_binding.allow_clear
              end
            end

            # Todo move to supper call we can have tool tips in other fields too
            unless a_binding.tooltip.nil? or a_binding.tooltip.empty?
              @casper_binding[:hint][:expression] = a_binding.tooltip
              a_generator.declare_expression_entities(a_binding.tooltip)
            end

          end

        end
      end
    end
  end
end