# -*- coding: utf-8 -*-
require 'rspec'
require File.expand_path(File.dirname(__FILE__) + "/../excel_model")

describe ExcelModel, "instance when it " do
  before do
    @excel = ExcelModel::Base.new('test')
  end
  
  it "should first sheet name is sample_sheet" do
    @excel.book.sheets[0].name.should eql "sample_sheet"
  end

end