require 'roo'

class Table < Array

  def initialize(path)
    workbook = Roo::Spreadsheet.open path
    worksheets = workbook.sheets
    worksheets.each do |worksheet|
      num_rows = 0
      workbook.sheet(worksheet).each_row_streaming do |row|
        row_cells = row.map { |cell| cell.value}
        flag = true
        row_cells.each do |s|
          if s.is_a? String
            if s.downcase == "subtotal" || s.downcase == "total"
              flag = false
            end
          end
        end  
        if flag
          self.push(row_cells)
        end
        num_rows += 1
      end
    end
  end

  def method_missing(m, *args)
    text = m.to_s
    if text == "row"
      self[args[0]]
    else
      column_counter = 0
      array = Array.new
      self[0].each do |t|
        column_counter += 1
        if text == t.delete(" ")
          self.each do |t1|
            counter = 0
            t1.each do |t2|
              counter += 1
              if counter == column_counter && t2 != t
                array.push(t2)
              end
            end
          end
        end
      end
      array    
    end
  end     
end

table = Table.new('./Book (4).xlsx')
table2 = Table.new('./Book 1 (2).xlsx')

table.each do |t|
  puts "#{t}"
end

table2.each do |t|
  puts "#{t}"
end

p table[1][1]

p table.row(0)[2]

p table.Indeks

p table.Bodovi.sum

Table.class_eval do
  def + (obj)
    return (self - obj) + (obj - self)
  end 
end

p table + table2

p table - table2















