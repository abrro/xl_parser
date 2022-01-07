require 'roo'
require './xl_parser'

table = ExcelSpreadsheet.new('./Book.xlsx')

p table.table

p table.table[0][1]

# table.each do |cell|
#     p cell
# end

p table.row(1)
p table.row(1)[0]

p table["Stanje"]
p table["Stanje"][3]

p table.Stanje.sum
p table.Andrej.ri13318

table.use_sheet(1)

p table.table