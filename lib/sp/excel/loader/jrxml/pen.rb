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

        class Pen

          attr_accessor :line_width
          attr_accessor :line_style
          attr_accessor :line_color

          def initialize
            @line_width = 1.0
            @line_style = 'Solid'
            @line_color = '#000000'
          end

          def attributes
            rv = Hash.new
            rv['lineWidth'] = @line_width unless @line_width.nil?
            rv['lineStyle'] = @line_style unless @line_style.nil?
            rv['lineColor'] = @line_color unless @line_color.nil?
            return rv
          end

          def to_xml (a_node)
            Nokogiri::XML::Builder.with(a_node) do |xml|
              xml.pen(attributes)
            end
          end

        end

        class TopPen < Pen

          def to_xml (a_node)
            Nokogiri::XML::Builder.with(a_node) do |xml|
              xml.topPen(attributes)
            end
          end

        end

        class LeftPen < Pen

          def to_xml (a_node)
            Nokogiri::XML::Builder.with(a_node) do |xml|
              xml.leftPen(attributes)
            end
          end

        end

        class RightPen < Pen

          def to_xml (a_node)
            Nokogiri::XML::Builder.with(a_node) do |xml|
              xml.rightPen(attributes)
            end
          end

        end

        class BottomPen < Pen

          def to_xml (a_node)
            Nokogiri::XML::Builder.with(a_node) do |xml|
              xml.bottomPen(attributes)
            end
          end

        end

      end
    end
  end
end
