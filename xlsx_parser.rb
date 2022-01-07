require 'roo'
require 'roo-xls'
require './column'
require 'rubyXL'

class ExcelSpreadsheet 

    attr_accessor :table,:col_table

    include Enumerable

    def initialize(path)
        @file = Roo::Spreadsheet.open(path, {:expand_merged_ranges => true})
        
        @table = get_table(@file.default_sheet)
        @col_table = @table.transpose
    end

    def get_table(sheet)
        temp_table = @file.sheet(sheet).to_matrix.to_a

        #Clears table from subtotal and total keywords and clear rows
        temp_table.delete_if do |array|
            if (array.all? nil) || (array.to_s.include? "subtotal") || (array.to_s.include? "total")
                true
            end
        end

        return temp_table
    end

    def use_sheet(sheet)
        @file.default_sheet = @file.sheets[sheet]
        @table = get_table(@file.sheets[sheet])
        @col_table = @table.transpose
    end

    def each
        0.upto(@table.length - 1) do |row|
            0.upto(@table[row].length - 1) do |col|
                yield @table[row][col]
            end
        end
    end

    def row(n)
        return @table[n]
    end

    def [](column_name)
        return @col_table[ @table[0].index(column_name.to_s) ]
    end

    def method_missing(column_name)
        Column.new(@table, column_name.to_s, @col_table[@table[0].index(column_name.to_s)])
    end

    def add_tables(table_1, table_2)


    end

    def delete_tables(table_1, table_2)

        
    end

end