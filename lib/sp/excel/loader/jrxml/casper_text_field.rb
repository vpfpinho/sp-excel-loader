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

          def initialize (a_generator, a_expression)
            super(Array.new, nil, nil)
            @value_types    = Array.new
            @casper_binding = {} 
            @generator      = a_generator
            @binding        = @generator.bindings[a_expression] 
            
            @generator.declare_expression_entities(a_expression)
            @text_field_expression = a_expression

            unless @binding.nil?
              if @binding.respond_to? :presentation
                pattern = @binding.presentation.format
                if pattern != nil and not pattern.empty?
                  if pattern.include? '$P{' or pattern.include? '$F{' or pattern.include? '$V{'
                    @pattern_expression = pattern
                  else
                    @pattern = pattern
                  end
                end
              end

              if @binding.respond_to? :editable 
                @casper_binding[:editable] = { is: @binding.editable }
              end

            end
            update_tooltip()
          end

          def update_tooltip ()
          
            if not @binding.nil?
              if @binding.respond_to? :tooltip
                tooltip = @binding.tooltip
                if not tooltip.nil? and not tooltip.empty?
                  @generator.declare_expression_entities(tooltip.strip)
                  @casper_binding[:hint] = { expression: tooltip.strip }
                end
              end
            end

          end

          def convert_type (a_value)
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

          def to_xml (a_node)
            puts '======================================================================================='
            puts "TextField   = '#{@text_field_expression}'" 
            puts "Pattern Exp = '#{@pattern_expression}'"    unless @pattern_expression.nil? or @pattern_expression.empty?
            puts "Pattern     = '#{@pattern}'"               unless @pattern.nil? or @pattern.empty?
            if  @casper_binding.size != 0
              puts 'casper-binding:'
              ap  @casper_binding
              @report_element.properties << Property.new('casper.binding', @casper_binding.to_json)
            end
            super(a_node)
          end

          def editable_conditional (a_value)
            @casper_binding[:editable] ||= {}
            @casper_binding[:editable][:conditionals] ||= {}
            @casper_binding[:editable][:conditionals][:is] = a_value
          end

          def disabled_conditional (a_value)
            @casper_binding[:editable] ||= {}
            @casper_binding[:editable][:conditionals] ||= {}
            @casper_binding[:editable][:conditionals][:disabled] = a_value
          end

          def locked_conditional (a_value)
            @casper_binding[:editable] ||= {}
            @casper_binding[:editable][:conditionals] ||= {}
            @casper_binding[:editable][:conditionals][:locked] = a_value
          end

          def enabled_conditional (a_value)
            @casper_binding[:editable] ||= {}
            @casper_binding[:editable][:conditionals] ||= {}
            @casper_binding[:editable][:conditionals][:enabled] = a_value
          end

          def style_expression (a_value)
            @casper_binding[:style] ||= {}
            @casper_binding[:style][:overload] ||= {}
            @casper_binding[:style][:overload][:condition] = a_value
          end
          
          def reload_if_changed (a_value)
            @casper_binding[:editable] ||= {}
            @casper_binding[:editable][:conditionals] |= {}
            @casper_binding[:editable][:conditionals][:reload] = a_value
          end
            
          def editable_expression (a_value)
            @casper_binding[:editable] ||= {}
            @casper_binding[:editable][:expression]  = a_value 
          end

        end

      end
    end
  end
end