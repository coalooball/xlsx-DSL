wb_path = File.join(XlsxFilesPath, 'xlsx_test1.xlsx')

describe OpenXML::SpreadsheetML::Workbook do

  OpenXML::SpreadsheetML::open wb_path do |wb|
    describe "Sheet#[]" do
      it "acquire value according index of cell" do  
        expect(wb['Sheet10']['D7']).to eql('D7')  
        expect(wb['some']['B5']).to eql('merge cell')  
        expect(wb['last']['E68']).to eql('E68cyan')  
      end
    end
    
    describe "Sheet#at_cell" do
      it "acquire value according index of cell" do  
        expect(wb['Sheet10'].at_cell('D7')).to eql('D7')  
        expect(wb['some'].at_cell('B5')).to eql('merge cell')  
      end
    end
  
    describe "Sheet#max_column" do
      it "acquire max index of column" do  
        expect(wb['Sheet3'].max_column).to eql('A')  
        expect(wb['Sheet4'].max_column).to eql('F')  
      end
    end
  
    describe "Sheet#max_row" do
      it "acquire max index of column" do  
        expect(wb['Sheet3'].max_row).to eql('1')  
        expect(wb['Sheet4'].max_row).to eql('27')  
      end
    end
  
    describe "Sheet#[]" do
      it "Call private retrieve_one_cell()" do  
        expect(wb['cell']['A1']).to eql('A1')  
        expect(wb['cell']['K74']).to eql('K74')  
      end
  
      it "Call private retrieve_one_row()" do  
        expect(wb['cell']['3']).to eq(['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11'])  
      end
  
      it "Call private retrieve_one_column()" do  
        expect(wb['some']['C']).to eq(['1', '2', '3', '4', '5'])  
        expect(wb['some']['B']).to eq(['10', '100', 'merge cell', 'merge cell', 'merge cell'])  
      end
  
      it "Call private retrieve_rows()" do  
        expect(wb['range']['1:2']).to eq([['1', '2', '3', '4'], ['5', '6', '7', '8']])  
        expect(wb['range']['3-4']).to eq([['9', '10', '11', '12'], ['13', '14', '15', '16']])  
      end
  
      it "Call private retrieve_columns()" do  
        expect(wb['range']['A-B']).to eq([['1', '5', '9', '13'], ['2', '6', '10', '14']])  
        expect(wb['range']['C:D']).to eq([['3', '7', '11', '15'], ['4', '8', '12', '16']])  
      end
  
      it "Call private retrieve_one_matrix()" do  
        expect(wb['range']['B2:C3']).to eq([['6', '7'], ['10', '11']])  
      end
    end
  end

end