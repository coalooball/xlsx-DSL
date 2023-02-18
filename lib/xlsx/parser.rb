module OpenXML
  module SpreadsheetML
    class Workbook
      def initialize xlsx_path
        @sheets = []
        @shared_strings = nil
        Zip::File.open(xlsx_path) do |zf|
          zf.each do |entry|
            content = entry.get_input_stream.read
            case entry.name
            when %r{^xl/workbook\.xml$}i
              workbook_parser content
            when %r{^xl/sharedStrings\.xml$}i
              shared_strings_parser content
            when %r{^xl/worksheets/sheet\d*\.xml$}i
              sheet_parser content
            end
          end
        end
        merge_sheets(@sheet_names) if @sheet_names
        merge_shared_strings(@shared_strings) if @shared_strings
        yield(self) if block_given?
      end
        
      class << self
        alias open new
      end

      private
  
      def workbook_parser content
        parser content
      end
  
      def shared_strings_parser content
        @shared_strings = SharedString.parser content || []
      end
  
      def sheet_parser content
        @sheets << Sheet.parser(content)
      end
    end
  end
end