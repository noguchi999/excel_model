# -*- coding: utf-8 -*-
require 'rspec'
require File.expand_path(File.dirname(__FILE__) + "/../excel_model")

describe ExcelModel, "instance when it " do
  before do
    @excel = ExcelModel::Base.new(File.expand_path("config/configuration.yml"))
  end
  
  it "should first sheet name is sample_sheet." do
    @excel.book.sheets[0].name.should eql "sample_sheet"
  end
  
  it "should can create instance with Hash augment." do
    configs = {development: {file_path: "test.xls", title_range: "B6:F6"}}
    @excel_by_hash_config  = ExcelModel::Base.new(configs)
    @excel_by_hash_config.book.sheets[0].name.should eql "sample_sheet"
  end
  
  it "should can create instance with String(not file) augment." do
    configs = <<-CONFIG
      !ruby/sym common: &common
          !ruby/sym file_path:             'test.xls'
          !ruby/sym title_range:           'B6:F6'

      !ruby/sym development:
        <<: *common

      !ruby/sym test:
        <<: *common

      !ruby/sym product:
        <<: *common
    CONFIG
    
    @excel_by_string_config  = ExcelModel::Base.new(configs)
    @excel_by_string_config.book.sheets[0].name.should eql "sample_sheet"
  end
  
  it "should raise runtime error when create instance with invalid type. " do
    lambda{ExcelModel::Base.new(1)}.should raise_error(RuntimeError, /Illegal augment type\..+?/)
  end
end