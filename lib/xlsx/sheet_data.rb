module OpenXML
  module SpreadsheetML
    SheetData = Struct.new(:rows)
    Row = Struct.new(:index, :columns)
    Cell = Struct.new(:type, :v, :formula)

    class Cell
      def value
        return $shared_strings[self.v.to_i] if self.type == 's'
        self.v
      end
    end
  end
end