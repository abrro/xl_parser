require 'roo'
require './xlsx_parser'

table1 = ExcelSpreadsheet.new('./Table1.xlsx')

p table1.table

p table1.table[0][1]

# table.each do |cell|
#     p cell
# end

p table1.row(1)
p table1.row(1)[0]

p table1["Index"]
p table1["Bodovi"][2]

p table1.Ocena.sum
p table1.Index.ri13318

table2 = ExcelSpreadsheet.new('./Table2.xlsx')

#table1.add_tables(table1.table, 0 , table2.table)

#table1.sub_tables(table1.original_table, 0, table2.original_table)