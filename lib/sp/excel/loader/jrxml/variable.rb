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

        class Variable

          attr_accessor :name
          attr_accessor :java_class
          attr_accessor :calculation
          attr_accessor :reset_type
          attr_accessor :variable_expression
          attr_accessor :initial_value_expression
          attr_accessor :presentation

          def initialize (a_name, a_java_class = nil)
            @name         = a_name
            @java_class   = a_java_class
            @java_class ||= 'java.lang.String'
            @calculation  = 'System'
            @reset_type   = nil
            @variable_expression = nil
            @initial_value_expression = nil
            @presentation = nil
          end

          def attributes
            rv = Hash.new
            rv['name']        = @name
            rv['class']       = @java_class
            rv['calculation'] = @calculation
            rv['resetType']   = @reset_type unless @reset_type.nil? or @reset_type == 'None'
            rv['resetGroup']  = 'Group1' if @reset_type == 'Group'
            return rv
          end

          def to_xml (a_node)
            Nokogiri::XML::Builder.with(a_node) do |xml|
              xml.variable(attributes) {
                unless @variable_expression.nil?
                  xml.variableExpression {
                    xml.cdata @variable_expression
                  }
                end
                unless @initial_value_expression.nil?
                  xml.initialValueExpression {
                    xml.cdata @initial_value_expression
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
