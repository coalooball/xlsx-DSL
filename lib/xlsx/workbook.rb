module OpenXML
  module SpreadsheetML
    class Workbook
      attr_accessor :sheets, :theme, :styles, :calc_chain, :xml_maps

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
        sheets = new_sheets
        self
      end

      def merge_shared_strings shared_strings_reference
        sheets.each do |sheet|
          sheet.merge_shared_strings shared_strings_reference
        end
      end

      private

      def parser content
        @sheet_names = []
        doc = Nokogiri::XML content
        sheet_tags = doc.css('sheet')
        sheet_tags.each do |s|
          sheet = Sheet.new
          sheet.name = s[:name]
          sheet.sheetId = s[:sheetId]
          sheet.rid = s['r:id']
          @sheet_names << sheet
        end
      end
    end
  end
end