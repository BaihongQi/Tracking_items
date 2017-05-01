require 'spreadsheet'
book = Spreadsheet.open('Copy of Shipment 56.xls')
sheet1 = book.worksheet('Shipping list')
sum=0
sheet1.each do |row|
  if (not row[1].nil? and not row[5].nil?) or (not row[1].nil? and not row[6].nil?)
    sum+=1
    if not row[5].nil? and row[5].match(/[[:upper:]]+[[:digit:]]{6}/)
      puts "No:#{sum} Name: #{row[1]} Node: #{row[5].match(/[[:upper:]]+[[:digit:]]{6}/)[0]}"
      #puts sum, row[1],row[5]
    elsif not row[6].nil? and row[6].match(/[[:upper:]]+[[:digit:]]{6}/)
      puts "No:#{sum} Name: #{row[1]} Node: #{row[6].match(/[[:upper:]]+[[:digit:]]{6}/)[0]}"
      #puts sum, row[1],row[6]
    end
    #puts row[1],row[5]
  end
  #puts row.join(',') # looks like it calls "to_s" on each cell's Value
end
