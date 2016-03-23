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

        class Extension

          attr_accessor :properties

          def initialize ()
            @properties = nil
          end

        end

        class Editable < Extension

          attr_accessor :field_id

          def initialize (a_field_id)

            if a_field_id[0] != "$"
              raise ArgumentError, "Invalid field id: '#{a_field_id}'!"
            end

            @field_id   = a_field_id[3..a_field_id.length-2]
            @properties = []
            @properties << Property.new("epaper.casper.text.field.editable.field_name", @field_id)
            @properties << Property.new("epaper.casper.text.field.editable", "true")
          end

        end

        class ClientCombo < Editable

          attr_accessor :id

          def initialize (a_field_id, a_id, a_uri)
            super(a_field_id)
            @id = a_id
            @properties << Property.new("epaper.casper.text.field.load.uri"                                            , a_uri)
            @properties << Property.new("epaper.casper.text.field.attach"                                              , "drop-down_list")
            @properties << Property.new("epaper.casper.text.field.attach.drop-down_list.controller"                    , "client")
            @properties << Property.new("epaper.casper.text.field.attach.drop-down_list.controller.display"            , "[id,name]")
            @properties << Property.new("epaper.casper.text.field.attach.drop-down_list.field.id"                      , "id")
            @properties << Property.new("epaper.casper.text.field.attach.drop-down_list.field.name"                    , "name")
            @properties << Property.new("epaper.casper.text.field.attach.drop-down_list.controller.pick.first_if_empty", "false")
            # @properties << Property.new("epaper.casper.text.field.patch.name"                              , "journal_id")
            # @properties << Property.new("epaper.casper.text.field.patch.type",                            , "java.lang.String")

          end

        end

        class Checkbox < Editable

          def initialize (a_field_id)
            super(a_field_id)
            @properties = [
                            Property.new("epaper.casper.text.field.editable"           , "false"),
                            Property.new("epaper.casper.text.field.editable.field_name", @field_id)
                          ]
          end

        end

        class RadioButton < Editable

          def initialize (a_field_id)
            super(a_field_id)
            @properties = [
                            Property.new("epaper.casper.text.field.editable"           , "false"),
                            Property.new("epaper.casper.text.field.editable.field_name", @field_id)
                          ]
          end

        end

        class ReportExtension < Extension

          attr_accessor :styles

          def initialize (a_report_name)
            super()
            @properties = [
                            Property.new("epaper.casper.text.field.editable.style"                , "EditableTextField"),
                            Property.new("epaper.casper.text.field.editable.style.focused"        , "EditableFocusedTextField"),
                            Property.new("epaper.casper.text.field.editable.style.disabled"       , "EditableDisabledTextField"),
                            Property.new("epaper.casper.text.field.editable.style.focused.invalid", "EditableFocusedInvalidContentTextField"),
                            Property.new("epaper.casper.text.field.editable.style.invalid"        , "EditableTextFieldInvalidContent")
                          ]

            @styles = []
            default = Style.new("EditableTextField")
              default.mode      ="Opaque"
              default.forecolor ="#000000"
              default.backcolor ="#D2EAF0"

              #if "iva_cover_sheet" == a_report_name or "iva_annex_r" == a_report_name
              # # cover
              # default.box       = default_box("#F2F2F2")
              # default.box.bottom_pen.line_color="#4B95FA"
              #else
              # # all others
              # default.box       = default_box("#000000")
              #end

            @styles << default

            invalid = Style.new("EditableTextFieldInvalidContent")
              invalid.style ="EditableTextField"
              invalid.box   = bottom_box("#E44A2C")
            @styles << invalid

            focused = Style.new("EditableFocusedTextField")
              focused.mode    = "Opaque"
              focused.forecolor = "#808080"
              focused.backcolor = "#F7F2E1"

              #if "iva_cover_sheet" == a_report_name or "iva_annex_r" == a_report_name
              # # cover
              # focused.box       = default_box("#CCCCCC")
              # focused.box.bottom_pen.line_color="#009BB5"
              #else
              # # all others
              # focused.box       = default_box("#000000")
              # focused.box.bottom_pen.line_color="#009BB5"
              #end

            @styles << focused

            # TODO NOT TESTED
            focused_invalid = Style.new("EditableFocusedInvalidContentTextField")
              focused_invalid.style = "EditableFocusedTextField"
              focused_invalid.box   = bottom_box("#E44A2C")
            @styles << focused_invalid

            # TODO NOT TESTED
            disabled = Style.new("EditableDisabledTextField")
              disabled.mode      = "Opaque"
              disabled.forecolor = "#C7C7C7"
              disabled.backcolor = "#F2F2F2"
              disabled.box = bottom_box("#000000", 1, "Dashed")
            @styles << disabled

          end

          def default_box (a_line_color, a_line_width=1, a_line_style="Solid")
            box            = Box.new
            box.left_pen   = LeftPen.new
            box.top_pen    = TopPen.new
            box.right_pen  = RightPen.new
            box.bottom_pen = BottomPen.new
            pens = [ box.left_pen, box.top_pen, box.right_pen, box.bottom_pen ]
            pens.each do |pen|
              pen.line_width = a_line_width
              pen.line_style = a_line_style
              pen.line_color = a_line_color
            end
            box
          end

          def bottom_box (a_line_color, a_line_width=1, a_line_style="Solid")
            box                       = Box.new
            box.bottom_pen            = BottomPen.new
            box.bottom_pen.line_width = a_line_width
            box.bottom_pen.line_style = a_line_style
            box.bottom_pen.line_color = a_line_color
            box
          end

          def new_pen (a_line_width, a_line_style, a_line_color)

          end

        end

        class WidgetFactory

          def initialize (a_map)
            @fields_map = a_map
          end

          def new_for_field (a_id)
            if a_id.match(/^\$P{/) || a_id.match(/^\$F{/)
              editable = @fields_map.has_key?(a_id) && @fields_map[a_id].editable ? Editable.new(a_id) : nil
            else
              editable = nil
            end
            pattern  = @fields_map[a_id].presentation.format if @fields_map.has_key?(a_id) && @fields_map[a_id].presentation != ''
            if editable.nil?
              widget = TextField.new(a_properties = nil, a_pattern = pattern, a_pattern_expression = nil)
            else
              widget = TextField.new(a_properties = editable.properties, a_pattern = pattern, a_pattern_expression = nil)
            end
            widget
          end

          def new_combo(a_config)
            config = a_config.strip[3..a_config.strip.length-2].strip.split(',')
            config[0].strip!
            config[1].strip!
            editable = ClientCombo.new(a_field_id=config[1], a_id=config[0], a_uri="model://#{config[0]}")
            widget = TextField.new(a_properties = editable.properties, a_pattern = nil, a_pattern_expression = nil)
            { id: editable.id, widget: widget, field:config[1] }
          end

          def new_checkbox(a_config)
            # check box: $CB{<field_name>,<unchecked>,<check
            cb       = a_config[4..a_config.length-2].split(',')
            editable = Checkbox.new(a_id=cb[0])
            widget   = TextField.new(a_properties = editable.properties, a_pattern = nil, a_pattern_expression = nil)
            widget.text_field_expression = "IF(#{cb[0]}==#{cb[2]};\"X\";\"\")"
            { widget: widget, field: cb[0] }
          end

          def new_radio_button(a_config)
            # check box: $RB{<field_name>,<unchecked>,<check
            rb       = a_config[4..a_config.length-2].split(',')
            editable = RadioButton.new(a_id=rb[0])
            widget   = TextField.new(a_properties = editable.properties, a_pattern = nil, a_pattern_expression = nil)
            widget.text_field_expression = "IF(#{rb[0]}==#{rb[2]};\"X\";\"\")"
            { widget: widget, field: rb[0] }
          end

          def java_class (a_id)
            if @fields_map.has_key?(a_id)
              @fields_map[a_id].java_class
            elsif '$V{PAGE_NUMBER}' == a_id || '$V{CONTINUOUS_PAGE_NUMBER}' == a_id
              'java.lang.Integer'
            else
              'java.lang.String'
            end
          end

        end

      end
    end
  end
end
