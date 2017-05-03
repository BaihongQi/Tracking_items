require 'spreadsheet'
require './helpers'
require 'csv'

def mysql_query(connection,query)
  begin
    rs = connection.query(query)
  rescue Exception => e
    raise e
    raise e if /Mysql::Error: Duplicate entry/.match(e.to_s)
  end
end


CSV.open("shipment52_not_in_openstack.csv","w") do |csv1|
book = Spreadsheet.open('Shipment 52.xls')
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

    check_exsit="select * from items where code='#{node}' or old_peel_new='#{node}'"
    rs1 = mysql_query(connection, check_exsit)
    row_num1=0
    rs1.each do |row1|
      row_num1+=1
    end
    if row_num1!=0
      check_mounted="select * from items where (code='#{node}' or old_peel_new='#{node}') and noid IS NULL"
      rs2 = mysql_query(connection, check_mounted)
      row_num2=0
      rs2.each do |row2|
        row_num2+=1
        csv1 << [row2['code'],row2['old_peel_new']]
      end

      puts "row number #{row_num1}------------------------------------------------"
    else
      puts "not in database"
      #csv1 << [row[1],row[5],row[6]]
      insert="INSERT INTO items_copy(code, digstatus,scanimages) VALUES ('#{node}','shipped','#{row[8]}') ON DUPLICATE KEY UPDATE digstatus=VALUES(digstatus), scanimages=VALUES(scanimages)"
      mysql_query(connection, check_mounted)
    end
  end
end
Helpers.close_mysql_connection(connection)
end
