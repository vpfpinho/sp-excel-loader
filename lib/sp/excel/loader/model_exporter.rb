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

      class ModelExporter < WorkbookLoader

        attr_accessor  :model

        def initialize (a_file, a_typed_export)
          super(a_file)
          @model        = Hash.new
          @typed_export = a_typed_export
        end

        def read_model(a_sheet_name, a_table_name)
          read_model_with_typed_option(a_sheet_name, a_table_name, @typed_export)
        end

        def read_model_with_typed_option(a_sheet_name, a_table_name, a_typed_export)

          read_cell_names(a_sheet_name)
          col_names       = Hash.new
          header_columns  = Hash.new
          scalar_formulas = Hash.new
          formula_lines   = Array.new
          scalar_values   = Hash.new
          value_lines     = Array.new

          worksheet  = @workbook[a_sheet_name]
          ref        = nil

          parse_shared_formulas(worksheet)

          # Capture the columns names
          worksheet.generic_storage.each do |tbl|
            next unless tbl.is_a? RubyXL::Table and tbl.name == a_table_name

            ref      = RubyXL::Reference.new(tbl.ref)
            type_row = ref.row_range.first() - 1
            i        = ref.col_range.first()
            tbl.table_columns.each do |table_col|
              if a_typed_export
                col_names[i] = { 'name' => table_col.name, 'type' => get_column_type(worksheet, type_row, i) }
              else
                col_names[i] = table_col.name
              end
              i += 1
            end

            col_names.sort.map do |key, value|
              header_columns[RubyXL::Reference.new(ref.row_range.first(),ref.col_range.first() + key - 1)] = value
            end
          end

          # Build the formula and value arrays
          for row in ref.row_range.begin()+1..ref.row_range.end()
            formula   = Hash.new
            value     = Hash.new
            col_index = 1

            ref.col_range.each do |col|

              cell = worksheet[row][col]
              if a_typed_export
                column = col_names[col_index]['name']
              else
                column = col_names[col_index]
              end
              if cell
                key, expression = cell_expression(cell)
                if cell.formula
                  formula[column] = expression
                else
                  value[column] = expression unless expression.nil?
                end
              end
              col_index += 1
            end
            formula_lines << formula
            value_lines << value
          end

          # Read scalar values and formulas, from the rows that are not part of the lines table
          renum = Array.new
          renum =  (0..ref.row_range.begin()).to_a
          renum += (ref.row_range.end()+1..worksheet.dimension.ref.row_range.end()).to_a
          for idx in renum
            next if worksheet[idx].nil?
            next if worksheet[idx].cells.nil?
            worksheet[idx].cells.each do |cell|
              if cell
                key, expression = cell_expression(cell)

                if cell.formula
                  scalar_formulas[key] = a_typed_export ? get_typed_scalar(cell, expression, worksheet) : expression
                else
                  unless expression.nil?
                    scalar_values[key] = a_typed_export ? get_typed_scalar(cell, expression, worksheet) : expression
                  end
                end
              end
            end
          end

          @model = {
            'values'   => scalar_values,
            'formulas' => scalar_formulas,
            'lines'    => {
              'header'    => header_columns,
              'formulas'  => formula_lines,
              'values'    => value_lines
              },
            }
        end

        def cell_expression (a_cell)

          cell_reference = RubyXL::Reference.ind2ref(a_cell.row, a_cell.column)
          name           = @cellnames[cell_reference]
          formula        = read_formula_expression(a_cell)

          if formula != nil
            if name
              expression = "#{name}=#{formula}"
            else
              expression = "#{cell_reference}=#{formula}"
            end
          elsif a_cell.value
            if name
              if a_cell.is_date?
                expression = "#{name}=excel_date(#{a_cell.value})"
              else
                begin
                  Float(a_cell.value)
                  expression = "#{name}=#{a_cell.value}"
                rescue
                  expression = "#{name}=\"#{a_cell.value}\""
                end
              end
            else
              if a_cell.is_date?
                expression = "#{cell_reference}=excel_date(#{a_cell.value})"
              else
                begin
                  Float(a_cell.value)
                  expression = "#{cell_reference}=#{a_cell.value}"
                rescue
                  expression = "#{cell_reference}=\"#{a_cell.value}\""
                end
              end
            end
          end
          return cell_reference, expression
        end

        def read_cell_names (a_sheet_name)
          @cellnames = {}
          ref_regexp = a_sheet_name + '!\$*([A-Z]+)\$*(\d+)'
          @workbook.defined_names.each do |dn|
            next unless dn.local_sheet_id.nil?
            match = dn.reference.match(ref_regexp)
            if match and match.size == 3
              matched_name = match[1].to_s + match[2].to_s
              if @cellnames[matched_name]
                raise "**** Fatal error:\n     duplicate cellname for #{matched_name}: #{@cellnames[matched_name]} and #{dn.name}"
              end
              @cellnames[matched_name] = dn.name
            end
          end
        end

        def get_column_type (a_worksheet, a_row_idx, a_column)
          return 'TEXT' if a_worksheet[a_row_idx].nil? or a_worksheet[a_row_idx][a_column].nil? or a_worksheet[a_row_idx][a_column].value.nil?
          return  a_worksheet[a_row_idx][a_column].value
        end

        def get_typed_scalar (a_cell, a_expression, a_worksheet)

          type = get_type_from_comment(a_cell.row, a_cell.column, a_worksheet)
          unless type.nil?
            return { 'type' => type, 'value' => a_expression }
          end

          if a_cell.is_date?
            return { 'type' => 'DATE', 'value' => a_expression }
          end
          case a_cell.datatype
          when RubyXL::DataType::NUMBER, nil
            return { 'type' => 'DECIMAL', 'value' => a_expression }
          when RubyXL::DataType::BOOLEAN
            return { 'type' => 'BOOLEAN', 'value' => a_expression }
          when RubyXL::DataType::DATE # Only available in Office2010
            return { 'type' => 'DATE', 'value' => a_expression }
          end
          return { 'type' => 'TEXT', 'value' => a_expression }
        end


        def get_type_from_comment (a_row, a_col, a_worksheet)

          if a_worksheet.comments != nil && a_worksheet.comments.size > 0 && a_worksheet.comments[0].comment_list != nil

            a_worksheet.comments[0].comment_list.each do |comment|
              if comment.ref.col_range.begin == a_col && comment.ref.row_range.begin == a_row
                comment.text.to_s.lines.each do |text|
                  text.strip!
                  next if text == '' or text.nil?
                  idx = text.index(':')
                  next if idx.nil?
                  tag   = text[0..(idx-1)]
                  value = text[(idx+1)..-1]
                  next if tag.nil? or value.nil?
                  tag.strip!
                  value.strip!

                  if tag == 'type'
                    return value
                  end
                end
              end
            end
          end
          return nil
        end

      end
    end
  end
end