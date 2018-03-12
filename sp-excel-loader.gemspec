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
# encoding: utf-8
#
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'sp/excel/loader/version'

Gem::Specification.new do |spec|
  spec.name          = 'sp-excel-loader'
  spec.version       = Sp::Excel::Loader::VERSION
  spec.email         = ['vitor.@servicepartner.pt']
  spec.date          = '2012-10-17'
  spec.summary       = 'Excelloader'
  spec.description   = 'Extends RubyXL adding handling of excel tables and other conversion utilies'
  spec.authors       = ['Vitor Pinho']
  spec.files         = Dir.glob("lib/**/*") + Dir.glob("spec/**/*") + %w(LICENSE README.md Gemfile)
  spec.homepage      = 'https://github.com/vpfpinho/sp-excel-loader.git'
  spec.license       = 'AGPL 3.0'
  spec.require_paths = ['lib']

  spec.add_development_dependency 'bundler', '~> 1.6'
  spec.add_development_dependency 'rake'
  spec.add_development_dependency 'rspec'

  spec.add_dependency 'rubyXL', '3.3.23'
  spec.add_dependency 'json'
end
