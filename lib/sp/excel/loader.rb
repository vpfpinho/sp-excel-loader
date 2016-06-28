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
require 'json'
#require 'byebug'

require File.expand_path(File.join(File.dirname(__FILE__), 'loader', 'rubyxl_table_patch'))
require File.expand_path(File.join(File.dirname(__FILE__), 'loader', 'workbookloader'))
require File.expand_path(File.join(File.dirname(__FILE__), 'loader', 'model_exporter'))
require File.expand_path(File.join(File.dirname(__FILE__), 'loader', 'payrollexporter'))
require File.expand_path(File.join(File.dirname(__FILE__), 'loader', 'json_to_xlsx'))
require File.expand_path(File.join(File.dirname(__FILE__), 'loader', 'jrxml', 'excel_to_jrxml'))
require File.expand_path(File.join(File.dirname(__FILE__), 'loader', 'jrxml', 'style'))
require File.expand_path(File.join(File.dirname(__FILE__), 'loader', 'jrxml', 'pen'))
require File.expand_path(File.join(File.dirname(__FILE__), 'loader', 'jrxml', 'band'))
require File.expand_path(File.join(File.dirname(__FILE__), 'loader', 'jrxml', 'band_container'))
require File.expand_path(File.join(File.dirname(__FILE__), 'loader', 'jrxml', 'box'))
require File.expand_path(File.join(File.dirname(__FILE__), 'loader', 'jrxml', 'group'))
require File.expand_path(File.join(File.dirname(__FILE__), 'loader', 'jrxml', 'parameter'))
require File.expand_path(File.join(File.dirname(__FILE__), 'loader', 'jrxml', 'field'))
require File.expand_path(File.join(File.dirname(__FILE__), 'loader', 'jrxml', 'variable'))
require File.expand_path(File.join(File.dirname(__FILE__), 'loader', 'jrxml', 'static_text'))
require File.expand_path(File.join(File.dirname(__FILE__), 'loader', 'jrxml', 'text_field'))
require File.expand_path(File.join(File.dirname(__FILE__), 'loader', 'jrxml', 'image'))
require File.expand_path(File.join(File.dirname(__FILE__), 'loader', 'jrxml', 'client_combo_text_field'))
require File.expand_path(File.join(File.dirname(__FILE__), 'loader', 'jrxml', 'report_element'))
require File.expand_path(File.join(File.dirname(__FILE__), 'loader', 'jrxml', 'jasper'))
require File.expand_path(File.join(File.dirname(__FILE__), 'loader', 'jrxml', 'property'))
require File.expand_path(File.join(File.dirname(__FILE__), 'loader', 'jrxml', 'property_expression'))
require File.expand_path(File.join(File.dirname(__FILE__), 'loader', 'jrxml', 'extensions'))
require File.expand_path(File.join(File.dirname(__FILE__), 'loader', 'version'))

module Sp
  module Excel
    module Loader

      def self.read_excel (a_path_name)
        we = PayrollExporter.parse(a_path_name)
        we
      end

      def self.read_excel_n_save (a_path_name, a_json_save_folder)
        we = PayrollExporter.parse(a_path_name)
        we.write_json_tables(a_json_save_folder)
        we.write_json_files(a_json_save_folder)
        we
      end

    end
  end
end
