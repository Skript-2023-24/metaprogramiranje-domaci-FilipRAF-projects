require 'roo'
require 'spreadsheet'

def open_excel(excel_file)
  return open_xlsx(excel_file) if excel_file.end_with? ".xlsx"
  return open_xls(excel_file) if excel_file.end_with? ".xls"

  print("Pogresna ekstenzija!")
  nil
end

def open_xlsx(excel_file)
  Spreadsheet.client_encoding = 'UTF-8'
  workbook = Roo::Spreadsheet.open(excel_file, { :expand_merged_ranges => true })
  process_worksheets(workbook.sheets, workbook)
end

def open_xls(excel_file)
  Spreadsheet.client_encoding = 'UTF-8'
  book = Spreadsheet.open excel_file
  process_worksheets(book.worksheets)
end

def process_worksheets(worksheets, workbook = nil)
  dimensions, elements = [], []

  worksheets.each do |worksheet|
    sheet = workbook ? workbook.sheet(worksheet) : worksheet
    process_worksheet(sheet, dimensions, elements)
  end

  build_table(dimensions, elements)
end

def process_worksheet(worksheet, dimensions, elements)
  worksheet.each do |row|
    row_cells = row.compact.map(&:to_s)
    next if row_cells.any? { |cell| cell.include?("total") || cell.include?("subtotal") }

    table_columns = row_cells.size
    elements.concat(row_cells)

    dimensions.push(table_columns) if table_columns > 0
  end
end

def build_table(dimensions, elements)
  col_number = dimensions.first
  help_array, matrix, counter = [], [], 0

  elements.each do |element|
    help_array << element
    counter += 1

    if counter == col_number
      matrix << help_array.clone
      help_array.clear
      counter = 0
    end
  end

  Table.new(matrix)
end

def add_method(c, method_name, &block)
  c.define_method(method_name, &block)
end

class Table
  include Enumerable
  attr_accessor :table, :t_table

  def initialize(matrix)
    @table = matrix
    @t_table = tran
    get_columns
    get_row_index
  end

  def row(num)
    @table[num]
  end

  def each_cell(&block)
    @table.each(&block)
  end

  def tran
    @table.first.zip(*@table[1..-1])
  end

  def [](column_name)
    create_hash[column_name]
  end


  def to_s
    @table.map { |row| row.join("\t") }.join("\n")
  end

  def create_hash
    h = Hash.new { |hash, key| hash[key] = [] }
    @t_table.each_with_index do |col, index|
      key = @table.first[index]
      col.drop(1).each do |value|
        h[key] << value unless ["total", "subtotal"].any? { |word| value.to_s.downcase.include?(word) }
      end
    end
    h
  end

  def get_columns
    @table.first.each_with_index do |cell, index|
      method_name = cell.to_sym
      add_method(self.class, method_name) { @t_table[index] }
    end
  end

  def get_row_index
    i = 1
    @t_table[0].slice(1..-1).each do |val|
      row_table = @table
      ij = i
      add_method(Array, val.to_s) {
        return row_table[ij]
      }
      i += 1
    end
  end

  def +(other)
    return nil if @table.first != other.table.first
    Table.new([@table.first] + @table.drop(1) + other.table.drop(1))
  end

  def -(other)
    return nil if @table.first != other.table.first
    remaining_rows = @table.drop(1) - other.table.drop(1)
    Table.new([@table.first] + remaining_rows)
  end
end

class Integer
  def include?(str)
    false
  end
end

class Float
  def include?(str)
    false
  end
end

class Array
  def sum
    self.map(&:to_f).reduce(0, :+)
  end

  def avg
    return 0 if self.empty?
    self.sum / (self.size - 1)
  end

end

t = open_excel("SJDomaci1.xlsx")

puts "1.2D niz:"
print t.table

puts "\n2.Pristupanje preko t.row(1):"
print t.row(1)

puts "\n3.Ima Enumerable i each:"
t.each_cell { |cell| puts cell }

puts "\n5.[] ima pristup poljima:"
print t["DrugaKolona"]
print "\n"
print t["DrugaKolona"][1]
print "\n"
print t["DrugaKolona"][1] = 2552

puts "\n6.Pristup preko istoimenih metod:"
print t.TrecaKolona
print "\n"
print t.TrecaKolona.sum
print "\n"
print t.TrecaKolona.avg
print "\n"
print t.Index.ri1721

t2 = open_excel("SJDomaci2.xls")

puts "\n9.Sabiranje dve tabele:"
t3 = t + t2
print t3.table

puts "\n10.Oduzimanje dve tabele:"
t4 = t3 - t2
print t4.table