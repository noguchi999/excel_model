#-*- coding: utf-8 -*-
require 'rspec/core/rake_task'

desc "Run all spec with RCov when AutoTest results are all green"
RSpec::Core::RakeTask.new('rcov') do |t|
  t.rcov = true
  t.rcov_opts = ['--exclude', 'spec']
end