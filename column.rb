require 'roo'

class Column

    attr_accessor :table, :column, :name 

    def initialize(table, name, column)
        @table = table
        @name = name
        @column = column
    end

    def sum
        return column.inject(0){|sum,x| sum + x.to_i }
    end

    def method_missing(row_name, *args)
        @column.each_with_index do |value, index|
            if value == row_name.to_s
                return @table[index]
            end
        end
        return nil
    end

end