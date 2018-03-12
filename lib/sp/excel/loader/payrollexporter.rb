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

      class PayrollExporter < ModelExporter

        attr_accessor  :tables
        attr_accessor  :model

        def initialize (a_file, a_typed_export)
          super(a_file, a_typed_export)
          @tables          = Hash.new
          @pretty          = true
        end

        def read_untyped_table (a_worksheet, a_table, a_table_name)

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

        def read_see_table (a_table_name)
          tbl = []

          tbl_instances = self.send "#{a_table_name}"

          # Get all columns reading all getter methods from first instance of a_table_name
          columns = tbl_instances.first.class.instance_methods(false).select { |method| method.to_s[-1] != '=' }

          columns.each do |column|
            col_obj  = Hash.new
            col_data = Array.new

            tbl_instances.each do |line|
              column_value = line.send(column)
              is_numeric = column_value.class.in?([Fixnum, BigDecimal])

              # We are at the first line of table, so prepare the structure
              if col_obj.empty?
                col_obj['name'] = column.to_s
                col_obj['type'] = is_numeric ? 'number' : 'text'
                col_obj['data'] = col_data
              end

              col_data << (is_numeric ? column_value.to_f : column_value.to_s)
            end

            tbl << col_obj
          end

          tbl
        end

        def read_all_untyped_tables ()
          @workbook.worksheets.each do |ws|
            ws.generic_storage.each do |tbl|
              next unless tbl.is_a? RubyXL::Table
              next if tbl.name == 'LINES'
              read_untyped_table(ws, tbl, tbl.name)
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

        def write_json_tables (a_directory)
          a_directory = File.join(a_directory, 'tables')
          @tables.each do |name, table|
            write_json_file(a_directory, name, table)
          end
        end

        def write_typed_json_tables (a_directory)
          a_directory = File.join(a_directory, 'tables')
          @table_names.each do |name|
            next if name == 'LINES'
            next if name != 'TABELA_RETROATIVOS'

            write_json_file(a_directory, name, read_see_table(name))
          end
        end

        def export(a_directory)
          self.read_all_untyped_tables # Legacy code when all tables of excel are typed
          self.read_all_tables
          self.write_json_file(a_directory, 'model',       self.read_model_with_typed_option('PROCESSAMENTO', 'LINES', false))
          self.write_json_file(a_directory, 'model_typed', self.read_model_with_typed_option('PROCESSAMENTO', 'LINES', true))
          self.write_json_tables(a_directory) # Legacy code when all tables of excel are typed
          self.write_typed_json_tables(a_directory)
          self
        end
      end
    end
  end
end