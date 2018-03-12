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

        class Box

          attr_accessor :top_pen
          attr_accessor :bottom_pen
          attr_accessor :left_pen
          attr_accessor :right_pen
          attr_accessor :padding
          attr_accessor :top_padding
          attr_accessor :leftPadding
          attr_accessor :bottom_padding
          attr_accessor :right_padding

          def initialize
            @top_pen        = nil
            @bottom_pen     = nil
            @left_pen       = nil
            @right_pen      = nil
            @padding        = 1
            @top_padding    = nil
            @leftPadding    = nil
            @bottom_padding = nil
            @right_padding  = nil
          end

          def attributes
            rv = Hash.new
            rv['padding']       = @padding        unless @padding.nil?
            rv['topPadding']    = @top_padding    unless @top_padding.nil?
            rv['leftPadding']   = @leftPadding    unless @leftPadding.nil?
            rv['bottomPadding'] = @bottom_padding unless @bottom_padding.nil?
            rv['rightPadding']  = @right_padding  unless @right_padding.nil?
            return rv
          end

          def to_xml (a_node)
            Nokogiri::XML::Builder.with(a_node) do |xml|
              xml.box(attributes)
            end
            @top_pen.to_xml(a_node.children.last)    unless @top_pen.nil?
            @left_pen.to_xml(a_node.children.last)   unless @left_pen.nil?
            @bottom_pen.to_xml(a_node.children.last) unless @bottom_pen.nil?
            @right_pen.to_xml(a_node.children.last)  unless @right_pen.nil?
          end

        end

      end
    end
  end
end
