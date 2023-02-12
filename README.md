# xlsx-DSL
A gem using DSL for reading&amp;writing Excel (.xlsx).

## Usage
###### Require gem
```ruby
require 'xlsx/DSL' 
```
### DSL
###### Select one worksheet
```ruby
OpenXML::SpreadsheetML::open 'new.xlsx' do |wb|
  wb['Sheet']
  wb.sheets.first
end
```
###### Acquire cells' value
```ruby
OpenXML::SpreadsheetML::open 'new.xlsx' do |wb|
  wb['Sheet']['A1']                          # a value of one cell
  wb['Sheet']['A']                           # values of one column
  wb['Sheet']['1']                           # values of one row
  wb['Sheet']['1:2']; wb['Sheet']['1-2']     # values of rows
  wb['Sheet']['A:C']; wb['Sheet']['A-C']     # values of columns
  wb['Sheet']['B2:C4']; wb['Sheet']['B2-C4'] # values of a matrix
end
```

### OOP
###### Select one worksheet
```ruby
wb = OpenXML::SpreadsheetML::open 'new.xlsx' # Initialize one Excel
ws = wb['Sheet']
ws = wb.sheets.first
```
###### Acquire cells' value
```ruby
ws['A1']                 # a value of one cell
ws['A']                  # values of one column
ws['1']                  # values of one row
ws['1:2']; ws['1-2']     # values of rows
ws['A:C']; ws['A-C']     # values of columns
ws['B2:C4']; ws['B2-C4'] # values of a matrix
```