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


CSV.open("shipment52_mounted.csv","w") do |csv1|
book = Spreadsheet.open('Shipment 52.xls')
sheet1 = book.worksheet('Shipping list')
sum=0#total
sum2=0#not parent
sum3=0#parent
connection = Helpers.set_mysql_connection
sheet1.each do |row|
  #check if p number exsit
  if not row[1].nil? and not row[5].nil? and row[5].match(/[[:upper:]]+[[:digit:]]{6}/)
    sum+=1
    title=row[1]
    node=row[5].match(/[[:upper:]]+[[:digit:]]{6}/)[0]
    #check if it
    if row[8].nil?
      first=node[1,2]
      second=node[3,2]
      whole=node[1,6]
      result=File.file?("/diginit/work/peel/metadata/N/#{first}/#{second}/#{node}.xml")
      if result==true
        sum3+=1
        puts "parent #{sum} #{sum3} #{node}"
        #csv1 << [row[1],row[5]]
        #save into a csv about parent items Name
      end
    end
    if not row[8].nil? or result!=true
      sum2+=1
        if not row[4].nil?
          year=row[4]
          puts "not parent #{sum} #{sum2}"
          #puts "No:#{sum} Name: #{title} Year: #{year} Node: #{node}"
        else
          puts "not parent #{sum} #{sum2}"
          #puts "No:#{sum} Name: #{title} Node: #{node}"
        end

        check_exsit="select * from items where code='#{node}' or old_peel_new='#{node}'"
        #puts "#{sum2}: #{node}"
        rs1 = mysql_query(connection, check_exsit)
        row_num1=0
        rs1.each do |row1|
          row_num1+=1
        end
        if row_num1!=0
          check_mounted="select * from items where (code='#{node}' or old_peel_new='#{node}') and digstatus='mounted'"
          #check_mounted="select * from items where (code='#{node}' or old_peel_new='#{node}') and noid IS NULL"
          rs2 = mysql_query(connection, check_mounted)
          row_num2=0
          rs2.each do |row2|
            row_num2+=1
            #csv1 << [row2['code'],row2['old_peel_new']]
          end

          #puts "row number #{row_num1}------------------------------------------------"
        else
          #puts "not in database #{node}"
          #csv1 << [row[1],row[5]]
          if not row[8].nil?
            insert="INSERT INTO items(code, digstatus,scanimages) VALUES ('#{node}','shipped','#{row[8]}') ON DUPLICATE KEY UPDATE digstatus=VALUES(digstatus), scanimages=VALUES(scanimages)"
          else
            insert="INSERT INTO items(code, digstatus) VALUES ('#{node}','shipped') ON DUPLICATE KEY UPDATE digstatus=VALUES(digstatus)"
          end
          mysql_query(connection, insert)
        end
      end
    end
end
Helpers.close_mysql_connection(connection)
end
