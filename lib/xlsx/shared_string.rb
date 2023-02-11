module OpenXML
  module SpreadsheetML
    class SharedString
      def self.parser content
        values = []
        doc = Nokogiri::XML content
        t_tags = doc.css('t')
        t_tags.each do |t|
          values << t.text
        end
        values
      end
    end
  end
end