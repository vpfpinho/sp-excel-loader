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

require 'bigdecimal'
require 'date'
require File.expand_path(File.join(File.dirname(__FILE__), 'rubyxl_table_patch'))

module Sp
  module Excel
    module Loader

      class TableRow
        def self.factory(klass_name = nil)
          if Object.constants.include?(klass_name.capitalize.to_sym)
            Object.send(:remove_const, klass_name.capitalize.to_sym)
          end
          klass = Class.new(self)
          Object.const_set(klass_name.capitalize, klass) unless klass_name.nil?
          klass
        end

        def add_attr(a_name, a_value)
          a_name = a_name.tr(' ', '_')
          self.class.send(:attr_accessor, a_name)
          instance_variable_set("@#{a_name}", a_value)
        end
      end

      class WorkbookLoader < TableRow

        attr_accessor :workbook

        def initialize (a_file)
          @workbook        = RubyXL::Parser.parse(a_file)
          @cellnames       = Hash.new
          @shared_formulas = Hash.new
          @table_names     = []
        end

        def read_all_tables ()
          @workbook.worksheets.each do |ws|
            ws.generic_storage.each do |tbl|
              next unless tbl.is_a? RubyXL::Table
              read_typed_table(ws, tbl, tbl.name)
              @table_names << tbl.name
            end
          end
        end

        def write_typed_table (a_records, a_table_name, a_style_filter = nil, a_keep_formulas = false)
          ws, tbl, ref = find_table(a_table_name)

          header_row = ref.row_range.begin()
          dst_row    = header_row + 1
          type_row   = header_row - 1

          a_records.each do |record|

            ref.col_range.each do |col|

              datatype, value = type_n_value_toxls(ws[type_row][col].value, record.send(ws[header_row][col].value))

              # Add or create cell
              if ws[dst_row].nil? || ws[dst_row][col].nil?
                ws.add_cell(dst_row, col, value)
              else
                if a_keep_formulas then
                  ws[dst_row][col].change_contents(value)
                else
                  style_index = ws[dst_row][col].style_index
                  ws.delete_cell(dst_row, col)
                  ws.add_cell(dst_row, col, value)
                  ws[dst_row][col].style_index = style_index
                end
              end
              # Only change datatype for number, other values make excel bark ...
              if [RubyXL::DataType::NUMBER].include? datatype
                ws[dst_row][col].datatype = datatype
              end

              # Call formater hook
              unless a_style_filter.nil?
                a_style_filter.call(record, ws, dst_row, col)
              end

            end
            dst_row += 1
          end

          # Adjust the table size
          previous_last_row = ref.col_range.end()
          tbl.ref = RubyXL::Reference.ind2ref(ref.row_range.begin(),
                                              ref.col_range.begin()) + ":" +
                    RubyXL::Reference.ind2ref(ref.row_range.begin() + a_records.size(),
                                              ref.col_range.end())
          for row in dst_row..previous_last_row
            ws.delete_row(row)
          end
        end

        #
        # @brief Convert the database value to the Excel cell type and value
        #
        # @param a_type  The value in the types header row
        # @param a_value Value from the database
        #
        # @return XLS type and value
        #
        def type_n_value_toxls (a_type, a_value)

          if a_value.nil?
            datatype = RubyXL::DataType::RAW_STRING
          else
            case a_type
            when 'INTEGER', 'INTEGER_NULLABLE'
              datatype = RubyXL::DataType::NUMBER
              a_value  = a_value.to_i
            when 'MONEY', 'MONEY_NULLABLE', 'DECIMAL', 'DECIMAL_NULLABLE'
              datatype = RubyXL::DataType::NUMBER
              a_value  = a_value.to_f
            when 'TEXT', 'TEXT_NULLABLE'
              datatype = RubyXL::DataType::RAW_STRING
            when 'BOOLEAN', 'BOOLEAN_NULLABLE'
              datatype = RubyXL::DataType::BOOLEAN
              unless a_value.nil?
                a_value = a_value ? 'TRUE' : 'FALSE'
              end
            when 'DATE', 'DATE_NULLABLE'
              datatype = RubyXL::DataType::DATE
              a_value  = @workbook.date_to_num(Date.iso8601(a_value))
            when 'DATETIME', 'DATETIME_NULLABLE'
              datatype = RubyXL::DataType::DATE
              a_value  = @workbook.date_to_num(DateTime.parse(a_value))
            else
              datatype = RubyXL::DataType::RAW_STRING
            end
          end
          return datatype, a_value
        end

        def read_typed_table (a_worksheet, a_table, a_table_name)
          ref        = RubyXL::Reference.new(a_table.ref)
          header_row = ref.row_range.begin()
          type_row   = header_row - 1
          records    = Array.new

          klass = TableRow.factory a_table_name

          for row in ref.row_range.begin()+1..ref.row_range.end()
            record = klass.new

            ref.col_range.each do |col|
              cell = a_worksheet[row][col] unless a_worksheet[row].nil?
              unless cell.nil?
                unless a_worksheet[type_row].nil? || a_worksheet[type_row][col].nil? || a_worksheet[type_row][col].value.nil?
                  type = a_worksheet[type_row][col].value
                else
                  type = 'TEXT'
                end
                case type
                when 'TEXT', 'TEXT_NULLABLE'
                  value = cell.value.to_s
                when 'SQL', 'SQL_NULLABLE'
                  value = cell.value.to_s
                when 'INTEGER', 'INTEGER_NULLABLE'
                  value = cell.value.to_i
                when 'DECIMAL', 'MONEY', 'DECIMAL_NULLABLE', 'MONEY_NULLABLE'
                  value = BigDecimal.new(cell.value.to_s)
                when 'BOOLEAN', 'BOOLEAN_NULLABLE'
                  value = cell.value.to_i == 0 ? false : true
                when 'DATE', 'DATE_NULLABLE'
                  begin
                    if cell.is_date?
                      value = cell.value
                    else
                      value = Date.parse cell.value
                    end
                  rescue => e
                    if type == 'DATE_NULLABLE'
                      value = nil
                    else
                      puts "Error in #{a_worksheet.sheet_name}!#{RubyXL::Reference.ind2ref(row,col)} #{e.message}"
                    end
                  end
                when 'DATETIME', 'DATETIME_NULLABLE'
                  begin
                    if cell.is_date?
                      value = cell.value
                    else
                      value = Date.parse cell.value
                    end
                  rescue => e
                    if type == 'DATETIME_NULLABLE'
                      value = nil
                    else
                      puts "Error in #{a_worksheet.sheet_name}!#{RubyXL::Reference.ind2ref(row,col)} #{e.message}"
                    end
                  end
                else
                  value = cell.value.to_s
                end
              else
                value = nil
              end
              if value.kind_of?(String)
                value = value.gsub(/\A\u00A0+/, '').gsub(/\u00A0+\z/, '').strip
              end
              begin
                record.add_attr(a_worksheet[header_row][col].value.strip, value)
              rescue => e
                puts "Error in #{a_worksheet.sheet_name}!#{RubyXL::Reference.ind2ref(header_row,col)} #{e.message}"
              end
            end
            records << record
          end
          add_attr(a_table_name, records)
        end

        def find_table (a_table_name)
          @workbook.worksheets.each do |ws|
            ws.generic_storage.each do |tbl|
              next unless tbl.is_a? RubyXL::Table
              next unless tbl.name == a_table_name

              return ws, tbl, RubyXL::Reference.new(tbl.ref)
            end
          end
          raise "Table '#{a_table_name}' not found in the workbook"
        end

        def read_named_cells (a_sheet_name)
          cellnames = Hash.new
          ref_regexp = a_sheet_name + '!\$*([A-Z]+)\$*(\d+)'
          @workbook.defined_names.each do |dn|
            next unless dn.local_sheet_id.nil?
            match = dn.reference.match(ref_regexp)
            if match and match.size == 3
              matched_name = match[1].to_s + match[2].to_s
              if cellnames[matched_name]
                raise "**** Fatal error:\n     duplicate cellname for #{matched_name}: #{@cellnames[matched_name]} and #{dn.name}"
              end
              cellnames[dn.name] = matched_name
            end
          end
          cellnames
        end

        ####################################################################################################################
        #                                                                                                                  #
        #                                  Methods that write XLS table into to the database                               #
        #                                                                                                                  #
        ####################################################################################################################

        def export_table_to_pg (a_conn, a_schema, a_prefix, a_table_name)
          export_table_to_pg_with_othername(a_conn, a_schema, a_prefix, a_table_name, a_table_name)
        end

        def export_table_to_pg_with_othername (a_conn, a_schema, a_prefix, a_table_name, a_xls_table_name)
          table   = a_schema
          table ||= 'public'
          table  += '.'
          table  += a_prefix unless a_prefix.nil?
          table  += a_table_name
          export_table_to_pg_other(a_conn, table, a_xls_table_name)
        end

        def export_table_to_pg_other (a_conn, a_db_table_name, a_xls_table_name)

          ws, tbl, ref = find_table(a_xls_table_name)

          header_row = ref.row_range.begin()
          dst_row    = header_row + 1
          type_row   = header_row - 1

          column_names = Array.new
          column_type  = Hash.new
          columns      = Array.new
          ref.col_range.each do |col|

            next if ws[type_row].nil?
            next if ws[type_row][col].nil?
            next if ws[type_row][col].value == 'VOID'

            column_type[col] = ws[type_row][col].value
            columns         << col
            column_names    << ws[header_row][col].value
          end

          rows = Array.new
          for row in ref.row_range.begin()+1..ref.row_range.end()

            next if ws[row].nil?
            row_values = Array.new
            columns.each do |col|

              cell = ws[row][col]
              if cell.nil?
                value = 'NULL'
              elsif column_type[col] =~ /_NULLABLE$/ && 0 == cell.value.to_s.strip.size
                value = 'NULL'
              else
                case column_type[col]
                when 'TEXT', 'TEXT_NULLABLE'
                  value = '\'' + a_conn.escape_string(cell.value.to_s.strip) + '\''
                when 'SQL', 'SQL_NULLABLE'
                  value = cell.value.to_s.strip
                when 'INTEGER', 'INTEGER_NULLABLE'
                  value = cell.value.to_i
                when 'DECIMAL', 'DECIMAL_NULLABLE', 'MONEY', 'MONEY_NULLABLE'
                  value = BigDecimal.new(cell.value.to_s.strip)
                when 'BOOLEAN', 'BOOLEAN_NULLABLE'
                  value = cell.value.to_i == 0 ? 'false' : 'true'
                when 'DATE', 'DATE_NULLABLE'
                  begin
                    value = '\'' + cell.value.strftime('%Y-%m-%d') + '\''
                  rescue => e
                    puts "Error in table #{a_xls_table_name} #{RubyXL::Reference.ind2ref(row,col)} #{e.message} value=#{cell.value.to_s}"
                  end
                when 'DATETIME', 'DATETIME_NULLABLE'
                  begin
                    value = '\'' + cell.value.strftime('%Y-%m-%d %H:%M:%S') + '\''
                  rescue => e
                    puts "Error in table #{a_xls_table_name} #{RubyXL::Reference.ind2ref(row,col)} #{e.message} value=#{cell.value.to_s}"
                  end
                else
                  value = '\'' + a_conn.escape_string(cell.value.to_s.strip) + '\''
                end
              end

              row_values << value
            end
            rows <<  '(' + row_values.join(',') + ')'
          end
          a_conn.exec("INSERT INTO #{a_db_table_name} (#{column_names.join(',')}) VALUES #{rows.join(",\n")}")
        end

        def write (a_filename)
          @workbook.calculation_chain = nil
          @workbook.calc_pr.full_calc_on_load = true
          @workbook.write(a_filename)
        end

        ####################################################################################################################
        #                                                                                                                  #
        #                       Methods related to cloning of calculation tables that use templates                        #
        #                                                                                                                  #
        ####################################################################################################################


        #
        # @brief This method takes a set of parameters read from the database and writes each parameter to excel cell on
        #        the specified sheet. The target cell must be named the column's name and have *no* formula
        #
        # @param a_sheet_name Name of the sheet where the replacements should be made
        # @param a_scalars    Hash with the values read from the database
        #
        def replace_scalars_in_sheet(a_sheet_name, a_scalars)

          ws = @workbook[a_sheet_name]
          cellnames = read_named_cells(a_sheet_name)

          a_scalars.each do |name, value|
            cell_id = cellnames[name]
            unless cell_id.nil?
              col, row = RubyXL::Reference.ref2ind(cell_id)
              cell = ws[col][row]
              if cell.formula.nil?
                if cell.is_date?
                  if value.nil? or value.length == 0
                    value = nil
                  else
                    value = @workbook.date_to_num(DateTime.iso8601(value))
                  end
                elsif (cell.datatype.nil? or cell.datatype == RubyXL::DataType::NUMBER)
                  begin
                    value = value.to_f
                  rescue
                    # Not a number? let it pass as it is
                  end
                end
                cell.change_contents(value)
              end
            end
          end
        end

        def parse_shared_formulas (a_worksheet)
          @shared_formulas = Hash.new
          ref = a_worksheet.dimension.ref
          ref.row_range.each do |row|
            ref.col_range.each do |col|
              next if a_worksheet[row].nil?
              cell = a_worksheet[row][col]
              next if cell.nil?
              formula = cell.formula
              next if formula.nil?

              if formula.t == 'shared' and not formula.expression.nil? and formula.expression.length != 0
                @shared_formulas[formula.si] = formula.expression
              end
            end
          end
        end

        def read_formula_expression (a_cell)
          return nil if (a_cell.formula.nil?)
          return @shared_formulas[a_cell.formula.si] if a_cell.formula.t == 'shared'
          return a_cell.formula.expression
        end

        def clone_lines_table (a_records, a_table_name, a_lines, a_template_column, a_closed_column = nil)
          ws, tbl, ref = find_table(a_table_name)

          header_row = ref.row_range.begin()
          style_row  = header_row + a_records.size
          dst_row    = style_row  + 1
          type_row   = header_row - 1

          template_index = Hash.new
          a_records.each_with_index do |record, index|
            template = record.send(a_template_column.to_sym)
            template_index[template] = header_row + index + 1
          end

          parse_shared_formulas(ws)

          a_lines.each_with_index do |line, index|
            if template_index.has_key?(line[a_template_column]) == false
              puts "Template #{line[a_template_column]} of line #{index+1} does exist in the model"
              next
            end

            src_row = template_index[line[a_template_column]]
            closed = line[a_closed_column] == 't'

            ref.col_range.each do |col|

              datatype, value = type_n_value_toxls(ws[type_row][col].value, line[ws[header_row][col].value])

              # Copy formula if the line is open
              expression = read_formula_expression(ws[src_row][col])
              if closed == false and not expression.nil? and expression.length != 0
                ws.add_cell(dst_row, col, '', expression)
              else
                ws.add_cell(dst_row, col, value)
              end
              ws[dst_row][col].style_index = ws[style_row][col].style_index
            end
            ws.change_row_height(dst_row, ws.get_row_height(src_row))
            dst_row += 1
          end

          # Adjust the table size
          previous_last_row = ref.col_range.end()
          tbl.ref = RubyXL::Reference.ind2ref(ref.row_range.begin(),
                                              ref.col_range.begin()) + ":" +
                    RubyXL::Reference.ind2ref(ref.row_range.begin() + a_records.size() + a_lines.ntuples,
                                              ref.col_range.end())
        end

      end
    end
  end
end
