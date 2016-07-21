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

      class JsonToXlsx < WorkbookLoader

        attr_accessor :json_data
        attr_accessor :fields

        def initialize (a_excel_template, a_json)
          super(a_excel_template)
          @json_data = a_json
        end

        def convert_to_xls ()
          ws, tbl, ref = find_table('lines')

          # Detect optional report mode
          is_report = false
          ws[0][0].value.lines.each do |line|
            directive, value = line.split(':')
            if directive.strip == 'IsReport' and value.strip == 'true'
              is_report = true
            end
          end

          headers_idx = 0 .. ref.row_range.begin() - 1
          footers_idx = ref.row_range.end() + 1 .. ws.count - 1

          # Replace parameters in header and footer rows
          (headers_idx.to_a + footers_idx.to_a).each do |row|
            ref.col_range.each do |col|
              next if ws[row].nil?
              next if ws[row][col].nil?
              value = ws[row][col].value
              next if value.nil?

              value = value.to_s
              @json_data['data']['attributes'].each do |key, val|
                value = value.gsub('$P{' + key +'}', val.to_s)
              end
              value = value.gsub('$V{PAGE_NUMBER}', '1')

              unless @json_data['included'].nil? or @json_data['included'].size == 0
                @json_data['included'][0]['attributes'].each do |key,val|
                  value = value.gsub('$F{' + key +'}', val.to_s)
                end
              end
              ws[row][col].change_contents(value)
            end
          end

          # Collect mapped fields
          header_row = ref.row_range.begin()
          dst_row    = header_row + 1
          fields     = Hash.new

          ref.col_range.each do |col|
            cell = ws[dst_row][col]
            next if cell.nil?
            next if cell.value.nil?
            m = /\A\$F{(.+)}\z/.match cell.value.strip()
            next if m.nil?
            fields[col] = m[1]
          end

          # Make space for the expanded data table, shift the merged cells down
          unless @json_data['included'].nil? or @json_data['included'].size == 0
            row_cnt = @json_data['included'].size
            if row_cnt != 0
              (row_cnt - 1).times { ws.insert_row(dst_row + 1) }
            end
            ws.merged_cells.each do |cell|
              next unless cell.ref.row_range.min >= dst_row
              cell.ref.instance_variable_set(:"@row_range", Range.new(cell.ref.row_range.min + row_cnt - 1, cell.ref.row_range.max + row_cnt - 1))
            end
          end

          # In report mode empty the first row and column
          if is_report
            ws.change_row_height(0, 6)
            ws.change_column_width(0, 1)
            ws.each_with_index do |row, ridx|
              ws.delete_cell(ridx, 0)
            end
          end

          # Create the table rows
          if @json_data['included'].nil?
            ref.col_range.each do |col|
              ws[dst_row][col].change_contents('')
              ws[dst_row][col].style_index = ws[header_row + 1][col].style_index
            end
          else
            @json_data['included'].each do |line|
              fields.each do |col,field|

                value = line['attributes'][field]
                if value.nil?
                  value = 0
                end

                if ws[dst_row].nil? || ws[dst_row][col].nil?
                  ws.add_cell(dst_row, col, value)
                else
                  ws[dst_row][col].change_contents(value)
                end
                ws[dst_row][col].style_index = ws[header_row + 1][col].style_index
              end
              dst_row += 1
            end
            # Update the table size
            tbl.ref = RubyXL::Reference.ind2ref(ref.row_range.begin(),
                                                ref.col_range.begin()) + ":" +
                      RubyXL::Reference.ind2ref(ref.row_range.begin() + (dst_row - header_row - 1),
                                                ref.col_range.end())
          end

          # In report mode delete all worksheets except the one that contains the lines table and wipe comments
          if is_report
            @workbook.worksheets.delete_if {|sheet| sheet.sheet_name != ws.sheet_name}
            ws.comments = Array.new
          end

        end

        def workbook
          @workbook
        end

      end

    end
  end
end