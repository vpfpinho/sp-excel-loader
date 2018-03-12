# encoding: utf-8
#
# Copyright (c) 2011-2016 Cloudware S.A. All rights reserved.
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

        class ClientComboTextField < TextField

          def initialize (a_binding, a_generator)
            super(Array.new, a_binding.presentation.format, nil)

            @report_element.properties << Property.new('epaper.casper.text.field.editable'                                            , 'true')
            @report_element.properties << Property.new('epaper.casper.text.field.load.uri'                                            , a_binding.uri)
            @report_element.properties << Property.new('epaper.casper.text.field.attach'                                              , 'drop-down_list')
            @report_element.properties << Property.new('epaper.casper.text.field.attach.drop-down_list.controller'                    , 'client')
            if a_binding.cc_field_name[0] != '['
              @report_element.properties << Property.new('epaper.casper.text.field.attach.drop-down_list.controller.display'          , "[#{a_binding.cc_field_id},#{a_binding.cc_field_name}]")
            else
              @report_element.properties << Property.new('epaper.casper.text.field.attach.drop-down_list.controller.display'          , a_binding.cc_field_name)
            end
            @report_element.properties << Property.new('epaper.casper.text.field.attach.drop-down_list.field.id'                      , a_binding.cc_field_id)
            if a_binding.cc_field_name[0] == '['
              @report_element.properties << Property.new('epaper.casper.text.field.attach.drop-down_list.field.name'                  , a_binding.cc_field_name[1..-2])
            else
              @report_element.properties << Property.new('epaper.casper.text.field.attach.drop-down_list.field.name'                  , a_binding.cc_field_name)
            end
            @report_element.properties << Property.new('epaper.casper.text.field.attach.drop-down_list.controller.pick.first_if_empty', 'false')
            if a_binding.respond_to?(:cc_field_patch) and a_binding.cc_field_patch != ''
              @report_element.properties << Property.new('epaper.casper.text.field.patch.name'                                        ,  a_binding.cc_field_patch)
            else
              @report_element.properties << Property.new('epaper.casper.text.field.patch.name'                                        , a_binding.id[3..-2])
            end
            @report_element.properties << Property.new('epaper.casper.text.field.patch.type'                                          , a_binding.java_class)

            unless a_binding.tooltip.nil? or a_binding.tooltip.empty?
              @report_element.properties << PropertyExpression.new('epaper.casper.text.field.hint.expression', a_binding.tooltip)
              a_generator.declare_expression_entities(a_binding.tooltip)
            end

            if a_binding.respond_to?(:allow_clear) 
              unless a_binding.allow_clear.nil? or !a_binding.allow_clear
                @report_element.properties << Property.new('epaper.casper.text.field.attach.drop-down_list.controller.add.empty_line', a_binding.allow_clear)
              end
            end

          end

        end

      end
    end
  end
end