#-*- coding: utf-8 -*-
require 'yaml'
require 'win32ole'
require 'spreadsheet'
require File.expand_path('../lib/string', __FILE__)

module ExcelModel
  class Base
    attr_reader :configuration, :records, :book
        
    def initialize(_configuration, env=:development)
      env = env.to_sym
    
      if _configuration.kind_of?(String)
        if FileTest.file?(_configuration)
          @configuration ||= YAML.load_file(_configuration)[env]
        else
          @configuration ||= YAML.load(_configuration)[env]
        end
      elsif _configuration.kind_of?(Hash)
        @configuration ||= _configuration[env]
      else
        raise "Illegal augment type. #{_configuration}"
      end
      
      if WIN32OLE_TYPE.progids.grep(/Excel.Application/).empty?
        @book = SpreadSheetWorkBookModel.new(@configuration)
      else
        @book = WorkBookModel.new(@configuration)
      end
      
      @records = @book.sheets[0].records
    end
  end
  
  class WorkBookModel
    attr_reader :sheets, :name
    
    def initialize(options = {})
      excel = WIN32OLE::new("excel.Application")
      excel.Visible = false
      @sheets = []
      begin
        workbook = excel.Workbooks.Open(File.expand_path(options[:file_path].encode("Windows-31J")))
        @name = File.basename(options[:file_path])
        sheet_names = options[:sheet_names] || workbook.Sheets
        sheet_names.each do |sheet|
          worksheet = sheet.respond_to?(:ole_methods) ? sheet : workbook.Sheets.Item(sheet["sheet_name"])
          next if options[:exclude_sheet_names].inject([]){|result, h| result << h.values; result.flatten}.include?(worksheet.name)
          
          @sheets << WorkSheetModel.new(worksheet, options)
        end
      rescue
        $stderr.puts $!
      ensure
        excel.Quit
      end
    end
  end
  
  class WorkSheetModel
    attr_reader :name, :records
    
    def initialize(worksheet, options = {})
      @name    = worksheet.name
      @records = Records.new(worksheet, options)
    end
  end
  
  class Records < Array
    
    def initialize(worksheet, options = {})      
      worksheet.Select
      options[:data_range] ||= detect_data_range(worksheet, options[:title_range])
      worksheet.Range(options[:data_range]).Value.each do |record|
        self << Record.new(worksheet.Range(options[:title_range]).Value.flatten, record)
      end
    end
    
    private
      def detect_data_range(worksheet, title_range)
        start_column = title_range[/^[A-Za-z]+/]
        end_column   = title_range[/:[A-Za-z]+/].delete(':')
        base_row     = title_range[/\d+?:/].delete(':').to_i + 1
        used_row     = base_row

        while worksheet.Range("#{start_column}#{used_row}").Value
          used_row = used_row + 1
        end

        "#{start_column}#{base_row}:#{end_column}#{used_row - 1}"
      end
  end
    
  class Record
    def initialize(titles, record)
      raise "title numbers discrepancy with data records" unless titles.size == record.size
      
      title_no = 1
      titles.each_with_index do |title, i|
        next if title.nil?
        title = "title_#{title_no}" if title[/[^\x01-\x7E]/] #titleに全角文字が含まれている場合は、文字列をtitle_連番に変換する.
        
        record[i].gsub!(/\\/, '/') if record[i].respond_to?(:gsub)
        if record[i].class == Float && (record[i] - record[i].truncate).zero?
          record[i] = record[i].truncate.to_s
        else
          record[i] = record[i].to_s          
        end
        safe_value = record[i].sanitize

        eval %Q|@#{title.to_variable} = "#{safe_value}"|
        self.instance_eval "def #{title.to_variable}; @#{title.to_variable}; end"
        
        title_no += 1
      end
    end
  end
  
  class SpreadSheetWorkBookModel
    attr_reader :sheets, :name
    
    def initialize(options = {})
      @sheets = []
      begin
        workbook = Spreadsheet.open(File.expand_path(options[:file_path].encode("Windows-31J")), 'rb')
        name = File.basename(options[:file_path])
        sheet_names = options[:sheet_names] || workbook.worksheets
        sheet_names.each do |sheet|
          worksheet = sheet.class == Spreadsheet::Excel::Worksheet ? sheet : workbook.worksheet(sheet["sheet_name"])
          next if options[:exclude_sheet_names].inject([]){|result, h| result << h.values; result.flatten}.include?(worksheet.name)
          
          @sheets << SpreadSheetWorkSheetModel.new(worksheet, options)
        end
      rescue
        $stderr.puts $!
      end
    end
  end
  
  class SpreadSheetWorkSheetModel
    attr_reader :name, :records
    
    def initialize(worksheet, options = {})
      @name    = worksheet.name
      @records = SpreadSheetRecords.new(worksheet, options)
    end
  end
  
  class SpreadSheetRecords < Array
    
    def initialize(worksheet, options = {})      
      options[:data_range] ||= detect_data_range(worksheet, options[:title_range])
      title_row     = options[:title_range][/\d+?$/].to_i - 1
      start_of_rows = options[:data_range][/\d+?:/].delete(':').to_i - 1
      end_of_rows   = options[:data_range][/\d+?$/].to_i - 1
      
      start_of_rows.upto(end_of_rows).each do |pos|
        self << Record.new(worksheet.row(title_row), worksheet.row(pos))
      end
    end
    
    private
      def detect_data_range(worksheet, title_range)
        start_column = title_range[/^[A-Za-z]+/]
        end_column   = title_range[/:[A-Za-z]+/].delete(':')
        base_row     = title_range[/\d+?:/].delete(':').to_i + 1
        used_row     = base_row
        
        while worksheet.row(used_row)[start_column.to_excel_number]
          used_row = used_row + 1
        end
        
        #row and column index are 0 initilize on SpreadSheet it differences between Excel. be careful what you setup index.
        "#{start_column}#{base_row}:#{end_column}#{used_row}"
      end
  end
end

class NilClass

  def inject(init)
    []
  end
end