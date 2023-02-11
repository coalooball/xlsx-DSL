xlsx_test1_path = File.join(XlsxFilesPath, 'xlsx_test1.xlsx')

describe OpenXML::SpreadsheetML::Workbook do

  let(:xlsx_test1) { OpenXML::SpreadsheetML::open xlsx_test1_path }

  it "acquire value according index of cell" do  
    expect(xlsx_test1['Sheet10']['D7'].value).to eq('D7')  
    expect(xlsx_test1['some']['B5'].value).to eq('merge cell')  
    expect(xlsx_test1['last']['E68'].value).to eq('E68cyan')  
  end

end