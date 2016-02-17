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

          def initialize (a_excel_filename, a_fields_map)
            super(a_excel_filename)
            report_name = File.basename(a_excel_filename, '.xlsx')
            @report = JasperReport.new(report_name)
            @current_band      = nil
            @first_row_in_band = 0
            @band_type         = nil
            @v_scale           = 1
            @widget_factory    = WidgetFactory.new(a_fields_map)

            generate_styles()

            @px_width = @report.page_width - @report.left_margin - @report.right_margin

            parse_sheets()
            File.write(report_name + '.jrxml', @report.to_xml)

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
              case xf.alignment.horizontal
              when 'left'
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
              when 'bottom'
                style.v_text_align ='Bottom'
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
              a_pen.line_style = ''
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
                #byebug
                puts  "Color from theme #{a_xls_color.theme} #{a_xls_color.tint}"
                return '#c0c0c0'
              elsif a_xls_color.auto or a_xls_color.rgb.nil?
                return '#000000'
              else
                return '#' + a_xls_color.rgb[2..-1]
              end
            else
              return "#INDEXED TODO"
            end
          end

          def parse_sheets
            @workbook.worksheets.each do |ws|
              @worksheet = ws
              @raw_width = 0
              for col in (1 .. @worksheet.dimension.ref.col_range.end)
                @raw_width += @worksheet.get_column_width_raw(col)
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
              @current_band.properties = [ Property.new("epaper.casper.band.patch.op.add.attribute.name", "data_row_type") ]
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
            when /VScale:.+/i
              @v_scale = a_row_tag.split(':')[1].strip.to_f
            when /Query:.+/i
              @report.query_string = a_row_tag.split(':')[1].strip
            when /Id:.+/i
              @report.id = a_row_tag.split(':')[1].strip
            else
              @current_band = nil
              @band_type    = nil
            end

            if @current_band != nil && @worksheet.comments != nil && @worksheet.comments.size > 0 && @worksheet.comments[0].comment_list != nil

              @worksheet.comments[0].comment_list.each do |comment|
                if comment.ref.col_range.begin == 0 && comment.ref.row_range.begin == a_row
                  text = comment.text.to_s
                  if text.start_with? 'PE:'
                    if @current_band.print_when_expression.nil?
                      @current_band.print_when_expression = text[3..-1].strip
                    end
                  end
                end
              end
            end

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

                field = create_field(row[col_idx].value.to_s)
                field.report_element.x = x_for_column(col_idx)
                field.report_element.y = y_for_row(a_row_idx)
                field.report_element.width  = cell_width
                field.report_element.height = cell_height
                field.report_element.style  = 'style_' + (row[col_idx].style_index + 1).to_s

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

          def create_field (a_expression)

            if ! (m = /\A\$P{(.+)}\z/.match a_expression.strip).nil?

              f_id                     = a_expression.strip
              rv                       = @widget_factory.new_for_field(f_id)
              rv.text_field_expression = a_expression

              add_parameter(f_id, m[1])

            elsif ! (m = /\A\$F{(.+)}\z/.match a_expression.strip).nil?

              f_id                     = a_expression.strip
              rv                       = @widget_factory.new_for_field(f_id)
              rv.text_field_expression = a_expression

              add_field(f_id.strip, m[1])

            elsif ! (m = /\A\$V{(.+)}\z/.match a_expression.strip).nil?

              f_id                     = a_expression.strip
              rv                       = @widget_factory.new_for_field(f_id)
              rv.text_field_expression = a_expression

              add_variable(f_id, m[1])

            elsif ! (m = /\A\$C{(.+)}\z/.match a_expression.strip).nil?

              # combo
              combo = @widget_factory.new_combo(a_expression.strip)
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

            else

              if a_expression.match(/^\$SE{/)

                expression = a_expression.strip

                expression.scan(/\$[A-Z]{[a-z_0-9\-#]+}/) { |v|
                  f_id = (/\A\$[PFV]{(.+)}\z/.match v).to_s
                  if false == f_id.nil?
                    f_nm = f_id[3..f_id.length-2]
                    if f_id.match(/^\$P{/)
                      add_parameter(f_id, f_nm)
                    elsif f_id.match(/^\$F{/)
                      add_parameter(f_id, f_nm)
                    elsif f_id.match(/^\$V{/)
                      add_parameter(f_id, f_nm)
                    else
                      raise ArgumentError, "Don't know how to add '#{f_id}'!"
                    end
                  end
                }
                rv = TextField.new(a_properties = nil, a_pattern = nil, a_pattern_expression = nil)
                rv.text_field_expression = expression[4..expression.length-2]
              else
                if a_expression.strip.include? "$"
                  rv = TextField.new(a_properties = nil, a_pattern = nil, a_pattern_expression = nil)
                  rv.text_field_expression = a_expression
                else
                  rv = StaticText.new
                  rv.text = a_expression
                end
              end


            end
            return rv
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

          def x_for_column (a_col_idx)

            width = 0
            for idx in (1 .. a_col_idx - 1) do
              width += @worksheet.get_column_width_raw(idx)
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
                height += scale_y(@worksheet.get_row_height(row))
              end
            end

            @current_band.height = height
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

            return 1, 1, scale_x(@worksheet.get_column_width_raw(a_col_idx)), scale_y(@worksheet.get_row_height(a_row_idx))

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
