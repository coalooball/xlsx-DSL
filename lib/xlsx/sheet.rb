module OpenXML
  module SpreadsheetML
    Sheet = Struct.new(
      :dimension, :sheet_data, :merge_cells, :sheet_views,
      :name, :sheetId, :rid
    )
    MergeCell = Struct.new(:range)

    class Sheet
      def + sheet
        self.dimension = self.dimension || sheet.dimension
        self.sheet_data = self.sheet_data || sheet.sheet_data
        self.merge_cells = self.merge_cells || sheet.merge_cells
        self.name = self.name || sheet.name
        self.sheetId = self.sheetId || sheet.sheetId
        self.rid = self.rid || sheet.rid
        self
      end

      def [] index
        case index
        when /^[A-Z]+\d+$/
          retrieve_one_cell index
        when /^\d+$/
          retrieve_one_row index
        when /^[A-Z]+$/
          retrieve_one_column index
        else
          raise IndexError, 'Invalid index'
        end
      end

      def at_cell index
        retrieve_one_cell index
      end

      def self.parser content
        doc = Nokogiri::XML content

        # sheetData
        sheet_data = {}
        sheet_data_tags = doc.css('sheetData')
        cell_tags = sheet_data_tags.css('c')
        cell_tags.each do |cell|
          t = cell[:t] 
          v = cell.at_css('v')
          v = v.text if v
          f = cell.at_css('f')
          f = f.text if f
          sheet_data[cell[:r]] = Cell.new(t, v, f)
        end

        # mergeCells
        merge_cells = []
        merge_cells_tags = doc.css('mergeCell')
        merge_cells_tags.each do |mc|
          merge_cells << MergeCell.new(mc[:ref])
        end

        # dimension
        dimension_tag = doc.at_css('dimension')
        dimension = dimension_tag[:ref]

        Sheet.new(dimension, sheet_data, merge_cells)
      end

      def max_column
        /([A-Z]+)\d+$/.match(dimension)[1]
      end

      def max_row
        /[A-Z]+(\d+)$/.match(dimension)[1]
      end

      private

      def retrieve_one_cell index
        self.merge_cells.each do |mc|
          if mc.include?(index)
            index = mc.top_left
          end
        end
        self.sheet_data[index]
      end

      def retrieve_one_row index
        cells = []
        ('A'..max_column).each do |col|
          cells << retrieve_one_cell(col+index)
        end
        cells
      end

      def retrieve_one_column index
        cells = []
        (1..max_row.to_i).each do |row|
          cells << retrieve_one_cell(index+row.to_s)
        end
        cells
      end

    end

    class MergeCell
      def include? index
        /([A-Z]+)(\d+):([A-Z]+)(\d+)/.match(self.range)
        col_span = ($1..$3)
        row_span = ($2..$4)
        /([A-Z]+)(\d+)/.match(index)
        col = $1
        row = $2
        return true if col_span.include?(col) && row_span.include?(row)
        false
      end

      def top_left
        self.range.split(/:/)[0]
      end
    end
  end
end