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

require 'set'

module Sp
  module Excel
    module Loader
      module Jrxml

        class ExcelToJrxml < WorkbookLoader

          attr_reader   :report

          def initialize (a_excel_filename, a_fields_map = nil, a_enable_cb_or_rb_edition=false, write_jrxml = true)
            super(a_excel_filename)
            read_all_tables()
            report_name = File.basename(a_excel_filename, '.xlsx')
            @report = JasperReport.new(report_name)
            @current_band                = nil
            @first_row_in_band           = 0
            @band_type                   = nil
            @v_scale                     = 1
            @detail_cols_auto_height     = false
            @auto_float                  = false
            @auto_stretch                = false
            @band_split_type             = nil
            @basic_expressions           = false

            # If the field map is not supplied load aux tables from the same excel
            if a_fields_map.nil?
              a_fields_map = Hash.new

              # Load parameters config table if it exists
              if respond_to?('params_def') and not params_def.nil?
                params_def.each do |param|
                  param.presentation = Presentation.new(param.presentation)
                  a_fields_map[param.id] = param
                end
              end

              # Load fields config table if it exists
              if respond_to?('fields_def') and not fields_def.nil?
                fields_def.each do |field|
                  field.presentation = Presentation.new(field.presentation)
                  a_fields_map[field.id] = field
                end
              end

              # Load variable definition table if it exists
              if respond_to? ('variables_def') and not variables_def.nil?
                variables_def.each do |vdef|
                  next if vdef.name.nil? or vdef.name.empty?
                  variable = Variable.new(vdef.name)
                  variable.java_class               = vdef.java_class         unless vdef.java_class.nil? or vdef.java_class.empty?
                  variable.calculation              = vdef.calculation        unless vdef.calculation.nil? or vdef.calculation.empty?
                  variable.reset_type               = vdef.reset              unless vdef.reset.nil? or vdef.reset.empty?
                  variable.variable_expression      = vdef.expression         unless vdef.expression.nil? or vdef.expression.empty?
                  variable.initial_value_expression = vdef.initial_expression unless vdef.initial_expression.nil? or vdef.initial_expression.empty?
                  @report.variables[vdef.name] = variable
                end
              end

            end

            @widget_factory             = WidgetFactory.new(a_fields_map)
            @widget_factory.cb_editable = a_enable_cb_or_rb_edition
            @widget_factory.rb_editable = a_enable_cb_or_rb_edition

            generate_styles()

            @px_width = @report.page_width - @report.left_margin - @report.right_margin

            parse_sheets()
            if write_jrxml
              File.write(report_name + '.jrxml', @report.to_xml)
            end
          end

          def generate_styles

            (0 .. @workbook.cell_xfs.size - 1).each do |style_index|
              style = xf_to_style(style_index)
              @report.styles[style.name] = style
            end

          end

          def xf_to_style (a_style_index)

            # create a style
            style = Style.new('style_' + (a_style_index + 1).to_s)

            # grab cell format
            xf = @workbook.cell_xfs[a_style_index]

            # Format font
            if xf.apply_font == true
              xls_font = @workbook.fonts[xf.font_id]

              if xls_font.name.val == 'Arial'
                style.font_name = 'DejaVu Sans Condensed'
              else
                style.font_name = xls_font.name.val
              end

              unless xls_font.color.nil?
                if xls_font.color.rgb.nil?
                  style.forecolor = '#FFFFFF'
                else
                  style.forecolor = convert_color(xls_font.color)
                end
              end

              style.font_size = xls_font.sz.val unless xls_font.sz.nil?
              style.is_bold   = true unless xls_font.b.nil?
              style.is_italic = true unless xls_font.i.nil?
            end

            # background
            if xf.apply_fill == true
              xls_fill = @workbook.fills[xf.fill_id]
              if xls_fill.pattern_fill.pattern_type == 'solid'
                style.backcolor = convert_color(xls_fill.pattern_fill.fg_color)
              end
            end

            # borders
            if xf.apply_border == true
              xls_border = @workbook.borders[xf.border_id]

              if xls_border.outline != nil

                if xls_border.outline.style != nil
                  style.box ||= Box.new
                  style.box.left_pen  = LeftPen.new
                  style.box.top_pen   = TopPen.new
                  style.box.right_pen = RightPen.new
                  style.box.bottom    = BottomPen.new
                  apply_border_style(style.box.left_pen  , xls_border.outline)
                  apply_border_style(style.box.top_pen   , xls_border.outline)
                  apply_border_style(style.box.right_pen , xls_border.outline)
                  apply_border_style(style.box.bottom_pen, xls_border.outline)
                end

              else

                if xls_border.left != nil && xls_border.left.style != nil
                  style.box ||= Box.new
                  style.box.left_pen = LeftPen.new
                  apply_border_style(style.box.left_pen, xls_border.left)
                end

                if xls_border.top != nil && xls_border.top.style != nil
                  style.box ||= Box.new
                  style.box.top_pen = TopPen.new
                  apply_border_style(style.box.top_pen, xls_border.top)
                end

                if xls_border.right != nil && xls_border.right.style != nil
                  style.box ||= Box.new
                  style.box.right_pen = RightPen.new
                  apply_border_style(style.box.right_pen, xls_border.right)
                end

                if xls_border.bottom != nil && xls_border.bottom.style != nil
                  style.box ||= Box.new
                  style.box.bottom_pen = BottomPen.new
                  apply_border_style(style.box.bottom_pen, xls_border.bottom)
                end

              end
            end

            # Alignment
            if xf.apply_alignment

              #byebug if a_style_index == 111
              case xf.alignment.horizontal
              when 'left', nil
                style.h_text_align ='Left'
              when 'center'
                style.h_text_align ='Center'
              when 'right'
                style.h_text_align ='Right'
              end

              case xf.alignment.vertical
              when 'top'
                style.v_text_align ='Top'
              when 'center'
                style.v_text_align ='Middle'
              when 'bottom', nil
                style.v_text_align ='Bottom'
              end

              # rotation
              case xf.alignment.text_rotation
              when nil
                style.rotation = nil
              when 0
                style.rotation = 'None'
              when 90
                style.rotation = 'Left'
              when 180
                style.rotation = 'UpsideDown'
              when 270
                style.rotation = 'Right'
              end

            end

            return style

          end

          def apply_border_style (a_pen, a_xls_border_style)
            case a_xls_border_style.style
            when 'thin'
              a_pen.line_width = 0.5
              a_pen.line_style = 'Solid'
            when 'medium'
              a_pen.line_width = 1.0
              a_pen.line_style = 'Solid'
            when 'dashed'
              a_pen.line_width = 1.0
              a_pen.line_style = 'Dotted'
            when 'dotted'
              a_pen.line_width = 0.5
              a_pen.line_style = 'Dotted'
            when 'thick'
              a_pen.line_width = 2.0
              a_pen.line_style = 'Solid'
            when 'double'
              a_pen.line_width = 0.5
              a_pen.line_style = 'Double'
            when 'hair'
              a_pen.line_width = 0.25
              a_pen.line_style = 'Solid'
            when 'mediumDashed'
              a_pen.line_width = 1.0
              a_pen.line_style = 'Dashed'
            when 'dashDot'
              a_pen.line_width = 0.5
              a_pen.line_style = 'Dashed'
            when 'mediumDashDot'
              a_pen.line_width = 1.0
              a_pen.line_style = 'Dashed'
            when 'dashDotDot'
              a_pen.line_width = 0.5
              a_pen.line_style = 'Dotted'
            when 'slantDashDot'
              a_pen.line_width = 0.5
              a_pen.line_style = 'Dotted'
            else
              a_pen.line_width = 1.0
              a_pen.line_style = 'Solid'
            end
            a_pen.line_color = convert_color(a_xls_border_style.color)
          end

          def convert_color (a_xls_color)
            if a_xls_color.indexed.nil?
              if a_xls_color.theme != nil
                cs = @workbook.theme.a_theme_elements.a_clr_scheme
                case a_xls_color.theme
                when 0
                  return tint_theme_color(cs.a_lt1, a_xls_color.tint)
                when 1
                  return tint_theme_color(cs.a_dk1, a_xls_color.tint)
                when 2
                  return tint_theme_color(cs.a_lt2, a_xls_color.tint)
                when 3
                  return tint_theme_color(cs.a_dk2, a_xls_color.tint)
                when 4
                  return tint_theme_color(cs.a_accent1, a_xls_color.tint)
                when 5
                  return tint_theme_color(cs.a_accent2, a_xls_color.tint)
                when 6
                  return tint_theme_color(cs.a_accent3, a_xls_color.tint)
                when 7
                  return tint_theme_color(cs.a_accent4, a_xls_color.tint)
                when 8
                  return tint_theme_color(cs.a_accent5, a_xls_color.tint)
                when 9
                  return tint_theme_color(cs.a_accent6, a_xls_color.tint)
                else
                  return '#c0c0c0'
                end

              elsif a_xls_color.auto or a_xls_color.rgb.nil?
                return '#000000'
              else
                return '#' + a_xls_color.rgb[2..-1]
              end
            else
              return "#INDEXED TODO"
            end
          end

          def tint_theme_color (a_color, a_tint)
            color   = a_color.a_sys_clr.last_clr unless a_color.a_sys_clr.nil?
            color ||= a_color.a_srgb_clr.val
            r = color[0..1].to_i(16)
            g = color[2..3].to_i(16)
            b = color[4..5].to_i(16)
            unless a_tint.nil?
              if ( a_tint <  0 )
                a_tint = 1 + a_tint;
                r = r * a_tint
                g = g * a_tint
                b = b * a_tint
              else
                r = r + (a_tint * (255 - r))
                g = g + (a_tint * (255 - g))
                b = b + (a_tint * (255 - b))
              end
            end
            r = 255 if r > 255
            g = 255 if g > 255
            b = 255 if b > 255
            color = "#%02X%02X%02X" % [r, g, b]
            color
          end

          def parse_sheets
            @workbook.worksheets.each do |ws|
              @worksheet    = ws
              @raw_width    = 0
              @current_band = nil
              @band_type    = nil
              for col in (1 .. @worksheet.dimension.ref.col_range.end)
                @raw_width += get_column_width(@worksheet, col)
              end
              generate_bands()
            end
          end

          def generate_bands ()

            for row in @worksheet.dimension.ref.row_range
              next if @worksheet[row].nil?
              next if @worksheet[row][0].nil?
              row_tag = @worksheet[row][0].value
              next if row_tag.nil?

              if @band_type != row_tag
                adjust_band_height()
                process_row_tag(row, row_tag)
                @first_row_in_band = row
              end
              unless @current_band.nil?
                generate_band_content(row)
              end
            end

            adjust_band_height()
          end

          def process_row_tag (a_row, a_row_tag)

            case a_row_tag
            when 'BG:'
              @report.band_containers << Background.new
              @current_band = @report.band_containers.last.band
              @band_type = a_row_tag
            when 'TL:'
              @report.band_containers << Title.new
              @current_band = @report.band_containers.last.band
              @band_type = a_row_tag
            when 'PH:'
              @report.band_containers << PageHeader.new
              @current_band = @report.band_containers.last.band
              @band_type = a_row_tag
            when 'CH:'
              @report.band_containers << ColumnHeader.new
              @current_band = @report.band_containers.last.band
              @band_type = a_row_tag
            when /DT\d*/
              @current_band = Band.new
              if @report.detail.nil?
                @report.detail = Detail.new
                @report.band_containers << @report.detail
              end
              @report.detail.bands << @current_band
              @band_type = a_row_tag
            when 'CF:'
              @report.band_containers << ColumnFooter.new
              @current_band = @report.band_containers.last.band
              @band_type = a_row_tag
            when 'PF:'
              @report.band_containers << PageFooter.new
              @current_band = @report.band_containers.last.band
              @band_type = a_row_tag
            when 'LPF:'
              @report.band_containers << LastPageFooter.new
              @current_band = @report.band_containers.last.band
              @band_type = a_row_tag
            when 'SU:'
              @report.band_containers << Summary.new
              @current_band = @report.band_containers.last.band
              @band_type = a_row_tag
            when 'ND:'
              @report.band_containers << NoData.new
              @current_band = @report.band_containers.last.band
              @band_type = a_row_tag
            when /GH\d*:/
              @current_band = Band.new
              @report.group ||= Group.new
              @report.group.group_header.bands << @current_band
              @band_type = a_row_tag
            when /GF\d*:/
              @current_band = Band.new
              @report.group ||= Group.new
              @report.group.group_footer.bands << @current_band
              @band_type = a_row_tag
            when /Orientation:.+/i
              @report.orientation = a_row_tag.split(':')[1].strip
              @report.update_page_size()
              @px_width = @report.page_width - @report.left_margin - @report.right_margin
            when /Size:.+/i
              @report.paper_size = a_row_tag.split(':')[1].strip
              @report.update_page_size()
              @px_width = @report.page_width - @report.left_margin - @report.right_margin
            when /Report.isTitleStartNewPage:.+/i
              @report.is_title_new_page =  a_row_tag.split(':')[1].strip == 'true'
            when /Report.leftMargin:.+/i
              @report.left_margin =  a_row_tag.split(':')[1].strip.to_i
              @px_width = @report.page_width - @report.left_margin - @report.right_margin
            when /Report.rightMargin:.+/i
              @report.right_margin =  a_row_tag.split(':')[1].strip.to_i
              @px_width = @report.page_width - @report.left_margin - @report.right_margin
            when /Report.topMargin:.+/i
              @report.top_margin =  a_row_tag.split(':')[1].strip.to_i
            when /Report.bottomMargin:.+/i
              @report.bottom_margin =  a_row_tag.split(':')[1].strip.to_i
            when /VScale:.+/i
              @v_scale = a_row_tag.split(':')[1].strip.to_f
            when /Query:.+/i
              @report.query_string = a_row_tag.split(':')[1].strip
            when /Id:.+/i
              @report.id = a_row_tag.split(':')[1].strip
            when /Group.isStartNewPage:.+/i
              @report.group ||= Group.new
              @report.group.is_start_new_page = a_row_tag.split(':')[1].strip == 'true'
            when /Group.isReprintHeaderOnEachPage:.+/i
              @report.group ||= Group.new
              @report.group.is_reprint_header_on_each_page = a_row_tag.split(':')[1].strip == 'true'
            when /BasicExpressions:.+i/
              @widget_factory.basic_expressions = a_row_tag.split(':')[1].strip == 'true'
            else
              @current_band = nil
              @band_type    = nil
            end

            if @current_band != nil && @worksheet.comments != nil && @worksheet.comments.size > 0 && @worksheet.comments[0].comment_list != nil

              @worksheet.comments[0].comment_list.each do |comment|
                if comment.ref.col_range.begin == 0 && comment.ref.row_range.begin == a_row
                  comment.text.to_s.lines.each do |text|
                    text.strip!
                    next if text == ''
                    tag, value =  text.split(':')
                    next if value.nil? || tag.nil?
                    tag.strip!
                    value.strip!
                    if tag == 'PE' or tag == 'printWhenExpression'
                      if @current_band.print_when_expression.nil?
                        @current_band.print_when_expression = value
                      end
                    elsif tag == 'lineParentIdField'
                      @current_band.properties ||= Array.new
                      @current_band.properties  << Property.new("epaper.casper.band.patch.op.add.attribute.name", value)
                    elsif tag == 'AF' or tag == 'autoFloat'
                      @current_band.auto_float = to_b(value)
                    elsif tag == 'AS' or tag == 'autoStretch'
                      @current_band.auto_stretch = to_b(value)
                    elsif tag == 'splitType'
                      @current_band.split_type = value
                    elsif tag == 'stretchType'
                      @current_band.stretch_type = value
                    end
                  end
                end
              end
            end

          end

          def to_b (a_value)
            a_value.match(/(true|t|yes|y|1)$/i) != nil
          end

          def generate_band_content (a_row_idx)

            row = @worksheet[a_row_idx]

            max_cell_height = 0
            col_idx         = 1

            while col_idx < row.size do

              col_span, row_span, cell_width, cell_height = measure_cell(a_row_idx, col_idx)

              if cell_width != nil

                if row[col_idx].nil? || row[col_idx].style_index.nil?
                  col_idx += col_span
                  next
                end

                field = create_field(row[col_idx])
                field.report_element.x = x_for_column(col_idx)
                field.report_element.y = y_for_row(a_row_idx)
                field.report_element.width  = cell_width
                field.report_element.height = cell_height
                field.report_element.style  = 'style_' + (row[col_idx].style_index + 1).to_s


                if @current_band.stretch_type
                  field.report_element.stretch_type = @current_band.stretch_type
                end

                if @current_band.auto_float and field.report_element.y != 0
                  field.report_element.position_type = 'Float'
                end

                if @current_band.auto_stretch and field.respond_to?('is_stretch_with_overflow')
                  field.is_stretch_with_overflow = true
                end

                # overide here with field by field directives
                process_field_comments(a_row_idx, col_idx, field)


                # If the field is from a horizontally merged cell we need to check the right side border
                if col_span > 1
                  field.box ||= Box.new
                  xf = @workbook.cell_xfs[row[col_idx + col_span - 1].style_index]
                  if xf.apply_border
                    xls_border = @workbook.borders[xf.border_id]

                    if xls_border.right != nil && xls_border.right.style != nil
                      field.box ||= Box.new
                      field.box.right_pen = RightPen.new
                      apply_border_style(field.box.right_pen, xls_border.right)
                    end
                  end
                end

                # If the field is from a vertically merged cell we need to check the bottom side border
                if row_span > 1
                  field.box ||= Box.new
                  xf = @workbook.cell_xfs[@worksheet[a_row_idx + row_span - 1][col_idx].style_index]
                  if xf.apply_border
                    xls_border = @workbook.borders[xf.border_id]

                    if xls_border.bottom != nil && xls_border.bottom.style != nil
                      field.box ||= Box.new
                      field.box.bottom_pen = BottomPen.new
                      apply_border_style(field.box.bottom_pen, xls_border.bottom)
                    end
                  end
                end
                if field_has_graphics(field)
                  @current_band.children << field
                  @report.style_set.add(field.report_element.style)
                end
              end
              col_idx += col_span
            end

          end

          def field_has_graphics (a_field)
            text_empty = false
            has_border = false
            opaque     = false

            if a_field.instance_of?(StaticText)
              if a_field.text.nil? || a_field.text.length == 0
                text_empty = true
              end
            end

            if a_field.instance_of?(TextField)
              if a_field.text_field_expression.nil? || a_field.text_field_expression.length == 0
                text_empty = true
              end
            end

            if a_field.box != nil
              if a_field.box.right_pen  != nil ||
                 a_field.box.left_pen   != nil ||
                 a_field.box.top_pen    != nil ||
                 a_field.box.bottom_pen != nil
                 has_border = true
              end
            end

            style = @report.styles[a_field.report_element.style]
            if style != nil
              if style.box != nil
                if style.box.right_pen  != nil ||
                   style.box.left_pen   != nil ||
                   style.box.top_pen    != nil ||
                   style.box.bottom_pen != nil
                  has_border = true
                end
              end
              if style.mode != nil && style.mode == 'Opaque'
                opaque = true
              end
            end

            return true if opaque

            return true if has_border

            return true unless text_empty

            return false
          end

          def create_field (a_cell)

            expression = a_cell.value.to_s
            if ! (m = /\A\$P{([a-zA-Z0-9_\-#]+)}\z/.match expression.strip).nil?

              # parameter
              f_id                     = expression.strip
              rv                       = @widget_factory.new_for_field(f_id, self)
              rv.text_field_expression = expression

              add_parameter(f_id, m[1])

            elsif ! (m = /\A\$F{([a-zA-Z0-9_\-#]+)}\z/.match expression.strip).nil?

              # field
              f_id                     = expression.strip
              rv                       = @widget_factory.new_for_field(f_id, self)
              rv.text_field_expression = expression

              add_field(f_id.strip, m[1])

            elsif ! (m = /\A\$V{([a-zA-Z0-9_\-#]+)}\z/.match expression.strip).nil?

              # variable
              f_id                     = expression.strip
              rv                       = @widget_factory.new_for_field(f_id, self)
              rv.text_field_expression = expression

              add_variable(f_id, m[1])

            elsif ! (m = /\A\$C{(.+)}\z/.match expression.strip).nil?

              # combo
              combo = @widget_factory.new_combo(expression.strip)
              rv    = combo[:widget]
              f_id  = combo[:field]
              f_nm  = f_id[3..f_id.length-2]

              if f_id.match(/^\$P{/)
                add_parameter(f_id, f_nm)
              elsif combo[:field].match(/^\$F{/)
                add_field(f_id, f_nm)
              elsif combo[:field].match(/^\$V{/)
                add_variable(f_id, f_nm)
              else
                raise ArgumentError, "Don't know how to add '#{f_id}'!"
              end

              rv.text_field_expression = "TABLE_ITEM(\"#{combo[:id]}\";\"id\";#{f_id};\"name\")"

            elsif expression.match(/^\$CB{/)

              # checkbox
              checkbox = @widget_factory.new_checkbox(expression.strip)
              declare_expression_entities(expression.strip)
              rv = checkbox[:widget]

            elsif expression.match(/^\$RB{/)

              # radio button
              declare_expression_entities(expression.strip)
              radio_button = @widget_factory.new_radio_button(expression.strip)
              rv = radio_button[:widget]

            elsif expression.match(/^\$DE{/)

              declare_expression_entities(expression.strip)
              de = expression.strip.split(',')
              de[0] = de[0][4..de[0].length-1]
              de[1] = de[1][0..de[1].length-2]

              properties = [
                              Property.new("epaper.casper.text.field.editable", "false"),
                              Property.new("epaper.casper.text.field.editable.field_name", de[0][3..de[0].length-2])
                           ]

              rv = TextField.new(a_properties = properties, a_pattern = nil, a_pattern_expression = nil)
              rv.text_field_expression = de[1]

            elsif expression.match(/^\$SE{/)

              declare_expression_entities(expression.strip)
              expression = expression.strip
              rv = TextField.new(a_properties = nil, a_pattern = nil, a_pattern_expression = nil)
              rv.text_field_expression = expression[4..expression.length-2]

            elsif expression.match(/^\$I{/)

              rv = Image.new()

              # copy cell alignment to image
              style = @report.styles['style_' + (a_cell.style_index + 1).to_s]
              rv.v_align = style.v_text_align
              rv.h_align = style.h_text_align

              unless expression.nil?
                expression = expression.strip
                rv.image_expression = transform_expression(expression[3..expression.length-2])
              end

            elsif expression.include? '$P{' or expression.include? '$F{' or expression.include? '$V{'

              expression = transform_expression(expression)
              rv = TextField.new(a_properties = nil, a_pattern = nil, a_pattern_expression = nil)
              rv.text_field_expression = expression.strip

            else

              rv = StaticText.new
              rv.text = expression

            end

            return rv
          end

          def declare_expression_entities (a_expression)

            a_expression.scan(/\$[A-Z]{[a-z_0-9\-#]+}/) { |v|
              f_id = (/\A\$[PFV]{(.+)}\z/.match v).to_s
              if false == f_id.nil?
                f_nm = f_id[3..-2]
                if f_id.match(/^\$P{/)
                  add_parameter(f_id, f_nm)
                elsif f_id.match(/^\$F{/)
                  add_field(f_id, f_nm)
                elsif f_id.match(/^\$V{/)
                  add_variable(f_id, f_nm)
                else
                  raise ArgumentError, "Don't know how to add '#{f_id}'!"
                end
              end
            }
            nil
          end

          def process_field_comments (a_row, a_col, a_field)

            if @worksheet.comments != nil && @worksheet.comments.size > 0 && @worksheet.comments[0].comment_list != nil

              @worksheet.comments[0].comment_list.each do |comment|
                if comment.ref.col_range.begin == a_col && comment.ref.row_range.begin == a_row
                  comment.text.to_s.lines.each do |text|
                    text.strip!
                    next if text == ''
                    tag, value =  text.split(':')
                    tag.strip!
                    value.strip!

                    if tag == 'PE' or tag == 'printWhenExpression'
                      a_field.report_element.print_when_expression = value
                    elsif tag == 'AF' or tag == 'autoFloat'
                      a_field.report_element.position_type = to_b(value) ? 'Float' : 'FixRelativeToTop'
                    elsif tag == 'AS' or tag == 'autoStretch' and a_field.respond_to?(:is_stretch_with_overflow)
                      a_field.is_stretch_with_overflow = to_b(value)
                    elsif tag == 'ST' or tag == 'stretchType'
                      a_field.report_element.stretch_type = value
                    elsif tag == 'BN' or tag == 'blankIfNull' and a_field.respond_to?(:is_blank_when_null)
                      a_field.is_blank_when_null = to_b(value)
                    elsif tag == 'PT' or tag == 'pattern' and a_field.respond_to?(:pattern)
                      a_field.pattern = value
                    end
                  end

                end
              end
            end
          end

          def transform_expression (a_expression)
            matches = a_expression.split(/(\$[PVF]{[a-zA-Z0-9_]+})/)
            if matches.nil?
              return a_expression
            end
            terms = Array.new
            matches.each do |match|
              if match.length == 0
                next
              elsif match.start_with?('$P{')
                terms << match
                add_parameter(match[0], match[3..-2])
              elsif match.start_with?('$F{')
                terms << match
                add_field(match[0], match[3..-2])
              elsif match.start_with?('$V{')
                terms << match
                add_variable(match[0], match[3..-2])
              else
                terms << '"' + match + '"'
              end
            end
            terms.join(' + ')
          end

          def add_parameter (a_id, a_name)
            unless @report.parameters.has_key? a_name
              parameter = Parameter.new(a_name, @widget_factory.java_class(a_id))
              @report.parameters[a_name] = parameter
            end
          end

          def add_field (a_id, a_name)
            unless @report.fields.has_key? a_name
              field = Field.new(a_name, @widget_factory.java_class(a_id))
              @report.fields[a_name] = field
            end
          end

          def add_variable (a_id, a_name)
            if "PAGE_NUMBER" != a_name
              unless  @report.variables.has_key? a_name
                variable = Variable.new(a_name, @widget_factory.java_class(a_id))
                @report.variables[a_name] = variable
              end
            end
          end

          def get_column_width (a_worksheet, a_index)
            width   = a_worksheet.get_column_width_raw(a_index)
            width ||= RubyXL::ColumnRange::DEFAULT_WIDTH
            return width
          end


          def x_for_column (a_col_idx)

            width = 0
            for idx in (1 .. a_col_idx - 1) do
              width += get_column_width(@worksheet, idx)
            end
            return scale_x(width)

          end

          def y_for_row (a_row_idx)
            height = 0
            for idx in (@first_row_in_band .. a_row_idx - 1) do
              height += @worksheet.get_row_height(idx)
            end
            return scale_y(height)
          end

          def adjust_band_height ()

            return if @current_band.nil?

            height = 0
            for row in @worksheet.dimension.ref.row_range
              unless @worksheet[row].nil? or @worksheet[row][0].nil? or @worksheet[row][0].value.nil? or @worksheet[row][0].value != @band_type
                height += @worksheet.get_row_height(row)
              end
            end

            @current_band.height = scale_y(height)
          end

          def measure_cell (a_row_idx, a_col_idx)

            @worksheet.merged_cells.each do |merged_cell|

              col_span = merged_cell.ref.col_range.size
              row_span = merged_cell.ref.row_range.size

              if a_row_idx == merged_cell.ref.row_range.begin && a_col_idx == merged_cell.ref.col_range.begin

                cell_height = y_for_row(merged_cell.ref.row_range.end + 1) -  y_for_row(merged_cell.ref.row_range.begin)
                cell_width  = x_for_column(merged_cell.ref.col_range.end + 1) - x_for_column(merged_cell.ref.col_range.begin)

                return col_span, row_span, cell_width, cell_height

              elsif merged_cell.ref.row_range.include?(a_row_idx) and merged_cell.ref.col_range.include?(a_col_idx)

                # The cell is overlaped by a merged cell
                return col_span, row_span, nil, nil

              end
            end

            return 1, 1, scale_x(get_column_width(@worksheet, a_col_idx)), scale_y(@worksheet.get_row_height(a_row_idx))

          end

          def scale_x (a_width)
            return (a_width * @px_width / @raw_width).round
          end

          def scale_y (a_height)
            return (a_height * @v_scale).round
          end


        end # class ExcelToJrxml

      end
    end
  end
end
