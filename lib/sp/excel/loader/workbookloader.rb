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

require 'bigdecimal'
require 'date'
require File.expand_path(File.join(File.dirname(__FILE__), 'rubyxl_table_patch'))

module Sp
  module Excel
    module Loader

      class TableRow

        def add_attr(a_name, a_value)
          self.class.send(:attr_accessor, a_name)
          instance_variable_set("@#{a_name}", a_value)
        end

      end

      class WorkbookLoader < TableRow

        def initialize (a_file)
          @workbook = RubyXL::Parser.parse(a_file)
        end

        def read_all_tables ()
          @workbook.worksheets.each do |ws|
            ws.generic_storage.each do |tbl|
              next unless tbl.is_a? RubyXL::Table
              read_typed_table(ws, tbl, tbl.name)
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

              var_name = ws[header_row][col].value
              value    = record.send(var_name)

              if value.nil?
                datatype = RubyXL::DataType::RAW_STRING
              else
                case ws[type_row][col].value
                when 'INTEGER'
                  datatype = RubyXL::DataType::NUMBER
                when 'MONEY', 'DECIMAL'
                  datatype = RubyXL::DataType::NUMBER
                when 'TEXT'
                  datatype = RubyXL::DataType::RAW_STRING
                when 'BOOLEAN'
                  datatype = RubyXL::DataType::BOOLEAN
                when 'DATETIME', 'DATE'
                  datatype = RubyXL::DataType::DATE
                else
                  datatype = RubyXL::DataType::RAW_STRING
                end
              end

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
              ws[dst_row][col].datatype = datatype

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

        def read_typed_table (a_worksheet, a_table, a_table_name)
          ref        = RubyXL::Reference.new(a_table.ref)
          header_row = ref.row_range.begin()
          type_row   = header_row - 1
          records    = Array.new

          for row in ref.row_range.begin()+1..ref.row_range.end()
            record = TableRow.new

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
                  value = DateTime.rfc3339(cell.value.to_s).to_date
                when 'DATETIME', 'DATETIME_NULLABLE'
                  value = DateTime.rfc3339(cell.value.to_s)
                else
                  value = cell.value.to_s
                end
              else
                value = nil
              end
              record.add_attr(a_worksheet[header_row][col].value, value)
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
rows_ins = Array.new
          for row in ref.row_range.begin()+1..ref.row_range.end()

            next if ws[row].nil?
            row_values = Array.new
row_inserts = Array.new
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
                  value = '\'' + DateTime.rfc3339(cell.value.to_s).to_date + '\''
                when 'DATETIME', 'DATETIME_NULLABLE'
                  value = DateTime.rfc3339(cell.value.to_s)
                else
                  value = '\'' + a_conn.escape_string(cell.value.to_s.strip) + '\''
                end
              end

              row_values << value
row_inserts << value.to_s.sub(/^'/,"''").sub(/'$/,"''")
            end
            rows <<  '(' + row_values.join(',') + ')'
rows_ins <<  '(' + row_inserts.join(',') + ')'
          end
          a_conn.exec("INSERT INTO #{a_db_table_name} (#{column_names.join(',')}) VALUES #{rows.join(",\n")}")

if ['vat_codes', 'vat_codes_conditions'].include? a_xls_table_name
puts %Q[EXECUTE 'INSERT INTO #{a_db_table_name} (#{column_names.join(',')}) VALUES #{rows_ins.join(",\n")};'\n]
end
        end

        def export_table_to_pg (a_conn, a_schema, a_prefix, a_table_name)
          table   = a_schema
          table ||= 'public'
          table  += '.'
          table  += a_prefix unless a_prefix.nil?
          table  += a_table_name
          export_table_to_pg_other(a_conn, table, a_table_name)
        end

        def write (a_filename)
          @workbook.calculation_chain = nil
          @workbook.calc_pr.full_calc_on_load = true
          @workbook.write(a_filename)
        end

      end

    end
  end
end
