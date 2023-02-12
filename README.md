# xlsx-DSL
A gem using DSL for reading&amp;writing Excel (.xlsx).

## Usage
###### Require gem
```ruby
require 'xlsx/DSL' 
```
### OOP
###### Open one Excel
```ruby
wb = OpenXML::SpreadsheetML::open 'new.xlsx'
```
###### Select one worksheet
```ruby
ws = wb['Sheet']
ws = wb.sheets.first
```
###### Acquire cells' value
```ruby
ws['A1']              # a value of one cell
ws['A']               # values of one column
ws['1']               # values of one row
ws['1:2']; ws['1-2']  # values of rows
ws['A:C']; ws['A-C']  # values of columns
```