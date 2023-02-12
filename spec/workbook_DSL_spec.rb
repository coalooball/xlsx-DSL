xlsx_test1_path = File.join(XlsxFilesPath, 'xlsx_test1.xlsx')

describe OpenXML::SpreadsheetML::Workbook do
  
  it "select worksheet according its name - Using DSL" do  
    OpenXML::SpreadsheetML::open xlsx_test1_path do |wb|
      expect(wb['some'].name).to eql('some')  
      expect(wb['last'].name).to eql('last')  
    end
  end
  
  it "select worksheet according index of sheets - Using DSL" do  
    OpenXML::SpreadsheetML::open xlsx_test1_path do |wb|
      expect(wb.sheets.first.name).to eql('Sheet10')  
      expect(wb.sheets[2].name).to eql('Sheet3')  
      expect(wb.sheets.last.name).to eql('last')  
    end
  end

end