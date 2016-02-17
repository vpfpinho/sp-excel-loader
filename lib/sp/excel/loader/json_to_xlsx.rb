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

          # Replace parameters in header rows
          (0..(ref.row_range.begin()-1)).each do |row|
            ref.col_range.each do |col|
              next if ws[row].nil?
              next if ws[row][col].nil?
              value = ws[row][col].value
              next if value.nil?

              value = value.to_s
              json_data['data']['attributes'].each do |key, val|
                value = value.gsub('$P{' + key +'}', val.to_s)
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
        end

        def workbook ; @workbook ; end

      end

    end
  end
end