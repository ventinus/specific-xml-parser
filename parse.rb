require 'roo'

show_empty = false
filename = "file"
start_row = 10
start_col = 6
end_col = 100
receipt_column = start_col - 2

sheet_titles = []

ARGV.each_with_index do |val, i|
  if val == '--filename'
    filename = ARGV[i + 1]
  end
end

xlsx = Roo::Spreadsheet.open("./#{filename}.xlsx")
all_sheets = xlsx.sheets

ARGV.each_with_index do |val, i|
  next if i % 2 == 1
  next_val = ARGV[i + 1]
  case val
  when '--start_row'
    start_row = next_val.to_i
  when '--sheet'
    sheet_titles << all_sheets[next_val.to_i - 1]
  when '--filename'
  when '--show-empty'
    show_empty = true
  else
    puts "#{val} is an unrecognized argument. Supported arguments are \"--start_row\", \"--sheet\", \"--filename\", \"--show-empty\""
  end
end

sheet_titles = all_sheets if sheet_titles.count.zero?

sheet_titles.each do |title|
  sheet = xlsx.sheet(title)
  delivery_row = sheet.row(start_row - 1)
  output_file = File.open("./output/#{title}.txt", "w")
  end_row = nil

  sheet.each_with_index do |row, i|
    next unless end_row.nil?
    if i > start_row && row[receipt_column].nil?
      end_row = i
    end
  end

  row_range = (start_row..end_row).to_a

  (start_col..end_col).to_a.each do |i|
    row_range.each do |j|
      next if delivery_row[i - 1].nil? || delivery_row[i - 1].class.name != 'Integer'

      if sheet.cell(j, i).zero?
        output_file.puts "[empty value]" if show_empty
      else
        output_file.puts "#{sheet.cell(j, receipt_column)}, #{delivery_row.at(i - 1)}, #{sheet.cell(j, i)}"
      end
    end
  end
end
