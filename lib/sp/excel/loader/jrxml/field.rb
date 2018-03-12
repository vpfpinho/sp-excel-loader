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

        class Field

          attr_accessor :name
          attr_accessor :java_class
          attr_accessor :description
          attr_accessor :default_value_expression
          attr_accessor :is_for_prompting

          def initialize (a_name, a_java_class = nil)
            @name         = a_name
            @java_class   = a_java_class
            @java_class ||= 'java.lang.String'
            @description  = nil
          end

          def attributes
            rv = Hash.new
            rv['name']  = @name
            rv['class'] = @java_class
            return rv
          end

          def to_xml (a_node)
            Nokogiri::XML::Builder.with(a_node) do |xml|
              xml.field(attributes) {
                unless @description.nil?
                  xml.fieldDescription {
                    xml.cdata @description
                  }
                end
              }
            end
          end

        end

      end
    end
  end
end
