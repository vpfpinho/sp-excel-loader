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

        class Style

          attr_accessor :name
          attr_accessor :style
          attr_accessor :mode
          attr_accessor :backcolor
          attr_accessor :forecolor
          attr_accessor :v_text_align
          attr_accessor :h_text_align
          attr_accessor :pattern
          attr_accessor :font_name
          attr_accessor :font_size
          attr_accessor :is_blank_when_null
          attr_accessor :is_bold
          attr_accessor :is_italic
          attr_accessor :pen
          attr_accessor :box
          attr_accessor :rotation

          def initialize (a_name)
            @name               = a_name
            @style              = nil
            @mode               = nil
            @backcolor          = nil
            @forecolor          = nil
            @v_text_align       = nil
            @h_text_align       = nil
            @pattern            = nil
            @font_size          = nil
            @font_name          = 'DejaVu Sans Condensed'
            @is_blank_when_null = true
            @is_bold            = nil
            @is_italic          = nil
            @pen                = nil
            @box                = nil
            @rotation           = nil
          end

          def to_xml (a_node)

            attrs = Hash.new
            attrs['name']            = @name
            attrs['style']           = @style              unless @style.nil?
            attrs['mode']            = @mode               unless @mode.nil?
            attrs['backcolor']       = @backcolor          unless @backcolor.nil?
            attrs['mode']            = 'Opaque'            unless @backcolor.nil?
            attrs['forecolor']       = @forecolor          unless @forecolor.nil?
            attrs['vTextAlign']      = @v_text_align       unless @v_text_align.nil?
            attrs['hTextAlign']      = @h_text_align       unless @h_text_align.nil?
            attrs['pattern']         = @pattern            unless @pattern.nil?
            attrs['fontSize']        = @font_size          unless @font_size.nil?
            attrs['fontName']        = @font_name          unless @font_size.nil? || @font_name == 'SansSerif'
            attrs['isBlankWhenNull'] = @is_blank_when_null unless @is_blank_when_null.nil?
            attrs['isBold']          = @is_bold            unless @is_bold.nil?
            attrs['isItalic']        = @is_italic          unless @is_italic.nil?
            attrs['rotation']        = @rotation           unless @rotation.nil?

            Nokogiri::XML::Builder.with(a_node) do |xml|
              xml.style(attrs)
            end
            unless @pen.nil?
              @pen.to_xml(a_node.children.last)
            end
            unless @box.nil?
              @box.to_xml(a_node.children.last)
            end
          end

        end # class Style

      end
    end
  end
end
