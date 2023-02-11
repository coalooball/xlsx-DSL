xlsx_test1_path = File.join(XlsxFilesPath, 'xlsx_test1.xlsx')

describe OpenXML::SpreadsheetML::Workbook do

  let(:xlsx_test1) { OpenXML::SpreadsheetML::open xlsx_test1_path }
  
  it "select worksheet according its name" do  
    expect(xlsx_test1['some'].name).to eq('some')  
    expect(xlsx_test1['last'].name).to eq('last')  
  end

  it "select worksheet according index of sheets" do  
    expect(xlsx_test1.sheets.first.name).to eq('Sheet10')  
    expect(xlsx_test1.sheets[2].name).to eq('Sheet3')  
    expect(xlsx_test1.sheets.last.name).to eq('last')  
  end

end