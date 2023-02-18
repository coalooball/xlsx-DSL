%w{
  zip
  nokogiri

  xlsx/parser
  xlsx/parser_helper
  xlsx/shared_string
  xlsx/sheet
  xlsx/sheet_data
  xlsx/workbook

  xlsx/DSL/version
}.each {|x| require x}