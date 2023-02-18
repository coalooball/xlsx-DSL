module OpenXML
  module SpreadsheetML
    def self.open xlsx_path
      Workbook.new xlsx_path
    end
  end
end