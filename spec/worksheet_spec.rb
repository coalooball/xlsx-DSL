xlsx_test1_path = File.join(XlsxFilesPath, 'xlsx_test1.xlsx')

describe OpenXML::SpreadsheetML::Workbook do

  let(:xlsx_test1) { OpenXML::SpreadsheetML::open xlsx_test1_path }

  describe "Sheet#[]" do
    it "acquire value according index of cell" do  
      expect(xlsx_test1['Sheet10']['D7']).to eql('D7')  
      expect(xlsx_test1['some']['B5']).to eql('merge cell')  
      expect(xlsx_test1['last']['E68']).to eql('E68cyan')  
    end
  end

  describe "Sheet#at_cell" do
    it "acquire value according index of cell" do  
      expect(xlsx_test1['Sheet10'].at_cell('D7')).to eql('D7')  
      expect(xlsx_test1['some'].at_cell('B5')).to eql('merge cell')  
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

  describe "Sheet#[]" do
    it "Call private retrieve_one_cell()" do  
      expect(xlsx_test1['cell']['A1']).to eql('A1')  
      expect(xlsx_test1['cell']['K74']).to eql('K74')  
    end

    it "Call private retrieve_one_row()" do  
      expect(xlsx_test1['cell']['3']).to eq(['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11'])  
    end

    it "Call private retrieve_one_column()" do  
      expect(xlsx_test1['some']['C']).to eq(['1', '2', '3', '4', '5'])  
      expect(xlsx_test1['some']['B']).to eq(['10', '100', 'merge cell', 'merge cell', 'merge cell'])  
    end

    it "Call private retrieve_rows()" do  
      expect(xlsx_test1['range']['1:2']).to eq([['1', '2', '3', '4'], ['5', '6', '7', '8']])  
      expect(xlsx_test1['range']['3-4']).to eq([['9', '10', '11', '12'], ['13', '14', '15', '16']])  
    end

    it "Call private retrieve_columns()" do  
      expect(xlsx_test1['range']['A-B']).to eq([['1', '5', '9', '13'], ['2', '6', '10', '14']])  
      expect(xlsx_test1['range']['C:D']).to eq([['3', '7', '11', '15'], ['4', '8', '12', '16']])  
    end
  end
end