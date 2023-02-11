module OpenXML
  module SpreadsheetML
    Workbook = Struct.new(
      :sheets, :theme, :styles, :calc_chain, 
      :xml_maps
    )

    class Workbook
      def [] index
        self.sheets.each do |sheet|
          return sheet if sheet.name == index
        end
        nil
      end
      
      def merge_sheets sheets
        new_sheets = []
        self.sheets.zip(sheets) do |x, y|
          new_sheets << x + y
        end
        self.sheets = new_sheets
        self
      end

      def self.parser content
        sheets = []
        doc = Nokogiri::XML content
        sheet_tags = doc.css('sheet')
        sheet_tags.each do |s|
          sheet = Sheet.new
          sheet.name = s[:name]
          sheet.sheetId = s[:sheetId]
          sheet.rid = s['r:id']
          sheets << sheet
        end
        Workbook.new(sheets = sheets)
      end
    end
  end
end