module OpenXML
  module SpreadsheetML
    class SharedString
      def initialize values # values is an array
        @values = values
      end

      def [] index
        @values[index]
      end

      def self.parser content
        values = []
        doc = Nokogiri::XML content
        t_tags = doc.css('t')
        t_tags.each do |t|
          values << t.text
        end
        SharedString.new values
      end
    end
  end
end