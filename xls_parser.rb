require 'roo'
require 'roo-xls'
require './column'
require 'spreadsheet'

class ExcelSpreadsheet 

    attr_accessor :table,:col_table,:original_table

    include Enumerable

    def initialize(path)
        @path = path
        @file = Roo::Spreadsheet.open(path)
        
        @original_table = @file.to_matrix.to_a
        @table = get_table(@file.default_sheet)
        @col_table = @table.transpose
    end

    def get_table(sheet)
        temp_table = @file.sheet(sheet).to_matrix.to_a

        #Clears table from subtotal and total keywords and clear rows
        temp_table.delete_if do |array|
            if (array.all? nil) || (array.to_s.downcase.include? "subtotal") || (array.to_s.downcase.include? "total")
                true
            end
        end

        #Missing expand merged cells option

        return temp_table
    end

    def use_sheet(sheet)
        @file.default_sheet = @file.sheets[sheet]
        @original_table = @file.sheet(sheet).to_matrix.to_a
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

    def add_tables(table_1, sheet, table_2)
        unless table_1[0] == table_2[0]
            puts "Error, table headers are not equal."
        end

        workbook = Spreadsheet.open @path
        sheet = workbook.sheet sheet

        1.upto(table_2.length - 1) do |row|
            0.upto(table_2[row].length - 1) do |col|
                sheet.rows[row + @file.last_row - 1][col + @file.first_column - 1] = table_2[row][col].to_s
            end
        end

        workbook.write @path
    end

    def sub_tables(table_1, sheet, table_2)

        unless table_1[0] == table_2[0]
            puts "Error, table headers are not equal."
        end

        same_elements = table_1 & table_2
        #remove header from list of same elements
        same_elements.delete_at(0)

        rows_to_remove = []

        same_elements.each do |arr|
            table_1.each_with_index do |row, index|
                if arr == row
                    p index
                    rows_to_remove.push(index)
                end
            end
        end

        workbook = Spreadsheet.open '/path/to/an/excel-file.xls'
        sheet = workbook.sheet sheet

        0.upto(rows_to_remove.length - 1) do |index|
            sheet.insert_row(rows_to_remove[index] + @file.first_row - 1, [])
        end

        workbook.write @path 
    end

end