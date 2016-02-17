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

        class Property

          attr_accessor :name
          attr_accessor :value

          def initialize (a_name, a_value)
            @name  = a_name
            @value = a_value
          end

          def attributes
            { name: @name, value:@value }
          end

          def to_xml(a_node)
            Nokogiri::XML::Builder.with(a_node) do |xml|
              xml.property(attributes)
            end
          end

        end

      end

    end
  end
end
