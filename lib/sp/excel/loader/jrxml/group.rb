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
      module Jrxml

        class Group

          attr_accessor :name
          attr_accessor :group_expression
          attr_accessor :group_header
          attr_accessor :group_footer
          attr_accessor :is_start_new_page
          attr_accessor :is_reprint_header_on_each_page

          def initialize (a_name = nil)
            @name = a_name || 'Group1'
            @group_expression  = '$F{data_row_type}'
            @is_start_new_page = nil
            @is_reprint_header_on_each_page = nil
            @group_header = GroupHeader.new
            @group_footer = GroupFooter.new
          end

          def attributes
            rv = Hash.new
            rv['name'] = @name
            rv['isStartNewPage'] = @is_start_new_page unless  @is_start_new_page  .nil?
            rv['isReprintHeaderOnEachPage'] = @is_reprint_header_on_each_page unless @is_reprint_header_on_each_page.nil?
            return rv
          end

          def to_xml (a_node)
            Nokogiri::XML::Builder.with(a_node) do |xml|
              xml.group(attributes)  {
                unless group_expression.nil?
                  xml.groupExpression {
                    xml.cdata @group_expression
                  }
                end
              }
            end
            @group_header.to_xml(a_node.children.last)
            @group_footer.to_xml(a_node.children.last)
          end

        end

      end
    end
  end
end
