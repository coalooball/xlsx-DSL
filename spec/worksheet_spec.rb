xlsx_test1_path = File.join(XlsxFilesPath, 'xlsx_test1.xlsx')

describe OpenXML::SpreadsheetML::Workbook do

  let(:xlsx_test1) { OpenXML::SpreadsheetML::open xlsx_test1_path }

  describe "Sheet#[]" do
    it "acquire value according index of cell" do  
      expect(xlsx_test1['Sheet10']['D7'].value).to eql('D7')  
      expect(xlsx_test1['some']['B5'].value).to eql('merge cell')  
      expect(xlsx_test1['last']['E68'].value).to eql('E68cyan')  
    end
  end

  describe "Sheet#at_cell" do
    it "acquire value according index of cell" do  
      expect(xlsx_test1['Sheet10'].at_cell('D7').value).to eql('D7')  
      expect(xlsx_test1['some'].at_cell('B5').value).to eql('merge cell')  
    end
  end

  describe "Sheet#max_column" do
    it "acquire max index of column" do  
      expect(xlsx_test1['Sheet3'].max_column).to eql('A')  
      expect(xlsx_test1['Sheet4'].max_column).to eql('F')  
    end
  end

  describe "Sheet#max_row" do
    it "acquire max index of column" do  
      expect(xlsx_test1['Sheet3'].max_row).to eql('1')  
      expect(xlsx_test1['Sheet4'].max_row).to eql('27')  
    end
  end
end