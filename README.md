# xlsx-DSL
A gem using DSL for reading&amp;writing Excel (.xlsx).

## Usage
###### Open one Excel
```ruby
wb = OpenXML::SpreadsheetML::open 'new.xlsx'
```
###### Select one worksheet
```ruby
ws = wb['Sheet']
ws = wb.sheets.first
```
###### Acquire cell's value
```ruby
ws['A1']
```