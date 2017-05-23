# coding: utf-8
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

  spec.add_dependency 'rubyXL'
  spec.add_dependency 'json'
end
