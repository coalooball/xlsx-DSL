module OpenXML
  module SpreadsheetML
    @@sheets = []
    @@shared_strings = nil

    class << self
      def open(xlsx_path)
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
        @@workbook.merge_sheets(@@sheets) if @@sheets
        yield(@@workbook) if block_given?
        @@workbook
      end

      private

      def workbook_parser content
        @@workbook = Workbook.parser content
      end

      def shared_strings_parser content
        $shared_strings = SharedString.parser content || []
      end

      def sheet_parser content
        @@sheets << Sheet.parser(content)
      end
    end
  end
end