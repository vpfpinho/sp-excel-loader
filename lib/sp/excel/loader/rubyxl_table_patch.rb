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

require 'rubyXL'
require 'rubyXL/objects/ooxml_object'

# Monkey patch to RubyXL
module RubyXL

  class SortCondition < OOXMLObject
    define_attribute(:ref           , :string, :required => true)
    define_element_name 'sortCondition'
  end

  class SortState < OOXMLObject
    define_attribute(:ref           , :string, :required => true)
    define_element_name 'sortState'
    define_child_node(RubyXL::SortCondition)
  end

  class TableColumn < OOXMLObject
    define_attribute(:id            , :int   , :required => true)
    define_attribute(:name          , :string, :required => true)
    define_attribute(:totalsRowShown, :int   , :default  => 0   )
    define_element_name 'tableColumn'
  end

  class TableColumns < OOXMLContainerObject
    define_attribute(:count, :int, :default => 0)
    define_child_node(RubyXL::TableColumn, :collection => true)
    define_element_name 'tableColumns'
  end

  class TableStyleInfo < OOXMLObject
    define_attribute(:name              , :string , :required => true)
    define_attribute(:showColumnStripes , :string , :default  => 0)
    define_attribute(:showFirstColumn   , :string , :default  => 0)
    define_attribute(:showLastColumn    , :string , :default  => 0)
    define_attribute(:showRowStripes    , :string , :default  => 0)
    define_element_name 'tableStyleInfo'
  end

  class Table < OOXMLTopLevelObject
    CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml'
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table'

    define_attribute(:id            , :int   , :required => true)
    define_attribute(:name          , :string, :required => true)
    define_attribute(:ref           , :string, :required => true)
    define_attribute(:displayName   , :string, :required => true)
    define_child_node(RubyXL::AutoFilter)
    define_child_node(RubyXL::SortState)
    define_child_node(RubyXL::TableColumns)
    define_child_node(RubyXL::TableStyleInfo)
    define_element_name 'table'
    set_namespaces('http://schemas.openxmlformats.org/spreadsheetml/2006/main' => '',
     'http://schemas.openxmlformats.org/officeDocument/2006/relationships' => 'r')

    def xlsx_path
      ROOT.join('xl', 'tables', "table#{file_index}.xml")
    end

  end

  class Tables < OOXMLContainerObject
    define_child_node(RubyXL::Table, :collection => true)
    define_element_name 'tables'
  end

  class Worksheet
    define_relationship(RubyXL::Table)
  end

end
