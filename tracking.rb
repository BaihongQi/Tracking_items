require 'spreadsheet'
require './helpers'

def mysql_query(connection,query)
  begin
    rs = connection.query(query)
  rescue Exception => e
    raise e
    raise e if /Mysql::Error: Duplicate entry/.match(e.to_s)
  end
end



book = Spreadsheet.open('Copy of Shipment 56.xls')
sheet1 = book.worksheet('Shipping list')
sum=0
connection = Helpers.set_mysql_connection
sheet1.each do |row|
  if (not row[1].nil? and not row[5].nil? and row[5].match(/[[:upper:]]+[[:digit:]]{6}/)) or (not row[1].nil? and not row[6].nil? and row[6].match(/[[:upper:]]+[[:digit:]]{6}/))
    sum+=1
    title=row[1]
    if not row[5].nil? and row[5].match(/[[:upper:]]+[[:digit:]]{6}/)
      node=row[5].match(/[[:upper:]]+[[:digit:]]{6}/)[0]
      if not row[4].nil?
        year=row[4]
        puts "No:#{sum} Name: #{title} Year: #{year} Node: #{node}"
      else
        puts "No:#{sum} Name: #{title} Node: #{node}"
      end
      #puts sum, row[1],row[5]
    elsif not row[6].nil? and row[6].match(/[[:upper:]]+[[:digit:]]{6}/)
      node=row[6].match(/[[:upper:]]+[[:digit:]]{6}/)[0]
      if not row[4].nil?
        year=row[4]
        puts "No:#{sum} Name: #{title} Year: #{row[4]} Node: #{node}"
      else
        puts "No:#{sum} Name: #{title} Node: #{node}"
      end
    end

    cmd="select * from items where code='#{node}'"
    puts cmd
    rs = mysql_query(connection, cmd)
    row_num=0
    rs.each do |row|
      row_num+=1
    end
    if row_num!=0
      puts "row number #{row_num}------------------------------------------------"
    else
      puts "not in database"
      #insert
    end
  end
end
Helpers.close_mysql_connection(connection)
