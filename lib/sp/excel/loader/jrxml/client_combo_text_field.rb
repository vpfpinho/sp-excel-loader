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

        class ClientComboTextField < TextField

          def initialize (a_binding)
            super(Array.new, a_binding.presentation.format, nil)

            @report_element.properties << Property.new('epaper.casper.text.field.editable'                                            , 'true')
            @report_element.properties << Property.new('epaper.casper.text.field.load.uri'                                            , a_binding.uri)
            @report_element.properties << Property.new('epaper.casper.text.field.attach'                                              , 'drop-down_list')
            @report_element.properties << Property.new('epaper.casper.text.field.attach.drop-down_list.controller'                    , 'client')
            @report_element.properties << Property.new('epaper.casper.text.field.attach.drop-down_list.controller.display'            , "[#{a_binding.cc_field_id},#{a_binding.cc_field_name}]")
            @report_element.properties << Property.new('epaper.casper.text.field.attach.drop-down_list.field.id'                      , a_binding.cc_field_id)
            @report_element.properties << Property.new('epaper.casper.text.field.attach.drop-down_list.field.name'                    , a_binding.cc_field_name)
            @report_element.properties << Property.new('epaper.casper.text.field.attach.drop-down_list.controller.pick.first_if_empty', 'false')
            @report_element.properties << Property.new('epaper.casper.text.field.patch.name'                                          , a_binding.id[3..-2])
            @report_element.properties << Property.new('epaper.casper.text.field.patch.type'                                          , a_binding.java_class)

          end

        end

      end
    end
  end
end