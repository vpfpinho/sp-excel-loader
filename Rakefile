require 'bundler/gem_tasks'
require 'rspec/core/rake_task'

RSpec::Core::RakeTask.new do |test|
  test.rspec_opts = "--color"
end

task :test => :spec

task :default => [:test]
