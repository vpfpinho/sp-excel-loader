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

      class PayrollExporter

        @workbook
        @tables
        @pretty
        @cellnames
        @model
        @shared_formulas

        attr_accessor  :tables
        attr_accessor  :model

        def initialize (a_file)
          @workbook        = RubyXL::Parser.parse(a_file)
          @tables          = Hash.new
          @pretty          = true
          @cellnames       = Hash.new
          @model           = Hash.new
          @shared_formulas = Hash.new
        end

        def read_table (a_worksheet, a_table, a_table_name)

          tbl = Array.new
          ref = RubyXL::Reference.new(a_table.ref)

          ref.col_range.each do |col|

            col_obj  = Hash.new
            col_data = Array.new

            type = 'number'
            for row in ref.row_range.begin()+1..ref.row_range.end()
              next if a_worksheet[row][col].nil?
              next if a_worksheet[row][col].value.nil?
              next if a_worksheet[row][col].value.is_a? Numeric
              next if a_worksheet[row][col].value.is_a? String and a_worksheet[row][col].value.length() == 0
              begin
                Float(a_worksheet[row][col].value)
              rescue
                type = 'text'
                break
              end
            end

            col_obj['name'] = a_worksheet[ref.row_range.begin()][col].value.to_s
            col_obj['type'] = type
            col_obj['data'] = col_data

            for row in ref.row_range.begin()+1..ref.row_range.end()
              if type == 'number'
                if a_worksheet[row][col].nil?
                  col_data << 0.0
                else
                  col_data << a_worksheet[row][col].value.to_f
                end
              else
                if a_worksheet[row][col].nil?
                  col_data << ''
                else
                  col_data << a_worksheet[row][col].value.to_s
                end
              end
            end

            tbl << col_obj
          end
          @tables[a_table.name] = tbl

        end

        def read_all_tables ()
          @workbook.worksheets.each do |ws|
            ws.generic_storage.each do |tbl|
              next unless tbl.is_a? RubyXL::Table
              next if tbl.name == 'LINES'
              read_table(ws, tbl, tbl.name)
            end
          end
        end

        def write_json_file (a_directory, a_name, a_object)
          FileUtils::mkdir_p(a_directory)
          File.open(File.join(a_directory, a_name + '.json'),"w") do |f|
            if @pretty
              f.write(JSON.pretty_generate(a_object))
            else
              f.write(a_object.to_json)
            end
          end
        end

        def read_cell_names (a_sheet_name)

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

        def write_json_files (a_directory)
          write_json_file(a_directory, 'model', @model)
        end

        def write_json_tables (a_directory)
          a_directory = File.join(a_directory, 'tables')
          @tables.each do |name, table|
            write_json_file(a_directory, name, table)
          end
        end

        def cell_expression (a_cell)

          cell_reference = RubyXL::Reference.ind2ref(a_cell.row, a_cell.column)
          name           = @cellnames[cell_reference]

          if a_cell.formula

            # Patch for shared formulas
            if a_cell.formula.t == 'shared' and a_cell.formula.expression = ""
              cr = RubyXL::Reference.new(a_cell.row, a_cell.row, a_cell.column, a_cell.column)
              for range in @shared_formulas.keys
                if range.cover?(cr)
                  a_cell.formula.expression = @shared_formulas[range]
                end
              end
            end

            if name
              expression = "#{name}=#{a_cell.formula.expression}"
            else
              expression = "#{cell_reference}=#{a_cell.formula.expression}"
            end
          elsif a_cell.value
            if name
              begin
                Float(a_cell.value)
                expression = "#{name}=#{a_cell.value}"
              rescue
                expression = "#{name}=\"#{a_cell.value}\""
              end
            else
              begin
                Float(a_cell.value)
                expression = "#{cell_reference}=#{a_cell.value}"
              rescue
                expression = "#{cell_reference}=\"#{a_cell.value}\""
              end
            end
          end
          return cell_reference, expression
        end

        def read_model (a_sheet_name, a_table_name)

          col_names       = Hash.new
          header_columns  = Hash.new
          model           = Hash.new
          scalar_formulas = Hash.new
          formula_lines   = Array.new
          scalar_values   = Hash.new
          value_lines     = Array.new

          worksheet  = @workbook[a_sheet_name]
          ref        = nil

          # Go hunting for shared formulas
          worksheet.each do |row|
            for col in (0..row.size())
              cell = row[col]
              if cell and cell.formula and cell.formula.ref and cell.formula.t = 'shared'
                @shared_formulas[cell.formula.ref] = cell.formula.expression
              end
            end
          end

          # Capture the columns names
          worksheet.generic_storage.each do |tbl|
            next unless tbl.is_a? RubyXL::Table and tbl.name == a_table_name

            ref = RubyXL::Reference.new(tbl.ref)
            i   = ref.col_range.first()
            tbl.table_columns.each do |table_col|
              col_names[i] = table_col.name
              i += 1
            end

            col_names.sort.map do |key,value|
              header_columns[RubyXL::Reference.new(ref.row_range.first(),ref.col_range.first() + key - 1)] = value
            end
          end

          # Build the formula and value arrays
          for row in ref.row_range.begin()+1..ref.row_range.end()
            formula   = Hash.new
            value     = Hash.new
            cell_ref  = RubyXL::Reference.new(row, row, ref.col_range.begin(), ref.col_range.end())
            col_index = 1

            ref.col_range.each do |col|

              cell           = worksheet[row][col]
              column         = col_names[col_index]

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
            worksheet[idx].cells.each do |cell|
              if cell
                key, expression = cell_expression(cell)
                if cell.formula
                  scalar_formulas[key] = expression
                else
                  scalar_values[key] = expression unless expression.nil?
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

        def self.parse (a_path_name)
          sheet_name = 'PROCESSAMENTO'
          we = PayrollExporter.new(a_path_name)
          we.read_all_tables
          we.read_cell_names(sheet_name)
          we.read_model(sheet_name, 'LINES')
          we
        end

      end

    end
  end
end