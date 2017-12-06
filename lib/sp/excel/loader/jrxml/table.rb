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

        class Table

          attr_accessor :relationship
          attr_accessor :column_header
          attr_accessor :detail
          attr_accessor :column_footer
          attr_accessor :no_data
          attr_accessor :groups

          def initialize ()
            @relationship  = nil
            @groups        = []
            @column_header = nil
            @detail        = Detail.new
            @column_footer = nil
            @no_data       = nil
          end

          def attributes
            rv = Hash.new
            rv['relationship'] = @relationship
            return rv
          end

          def to_xml (a_node)
            Nokogiri::XML::Builder.with(a_node) do |xml|
              xml.table(attributes)  {

              }
            end
            @groups.each do |group|
              group.to_xml(a_node.children.last)
            end
            @column_header.to_xml(a_node.children.last) unless @column_header.nil?
            @detail.to_xml(a_node.children.last) unless @detail.nil?
            @column_footer.to_xml(a_node.children.last) unless @column_footer.nil?
            @no_data.to_xml(a_node.children.last) unless @no_data.nil?
          end

        end

      end
    end
  end
end
