xlsx_test1_path = File.join(XlsxFilesPath, 'xlsx_test1.xlsx')

describe OpenXML::SpreadsheetML::Workbook do

  let(:xlsx_test1) { OpenXML::SpreadsheetML::open xlsx_test1_path }
  
  it "select worksheet according its name" do  
    expect(xlsx_test1['some'].name).to eql('some')  
    expect(xlsx_test1['last'].name).to eql('last')  
  end
  
  it "select worksheet according index of sheets" do  
    expect(xlsx_test1.sheets.first.name).to eql('Sheet10')  
    expect(xlsx_test1.sheets[2].name).to eql('Sheet3')  
    expect(xlsx_test1.sheets.last.name).to eql('last')  
  end
  
  it "test validity of Workbook.open" do 
    xlsx_open_method = OpenXML::SpreadsheetML::Workbook.open(xlsx_test1_path)
    expect(xlsx_open_method['some'].name).to eql('some')  
  end
end