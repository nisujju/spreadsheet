require 'spreadsheet'
Spreadsheet.client_encoding = 'UTF-8'
loan1 = Spreadsheet.open 'C:/Users/njha9/Desktop/JIRA/MINO-1374/Loan-apply.xls'
loan2 = Spreadsheet.open 'C:/Users/njha9/Desktop/JIRA/MINO-1374/Loan-ellen.xls'
sheet = loan1.worksheet 0
sheet2 = loan2.worksheet 0

 sheet.each do |row|
  a='y'	 
  sheet2.each do |row2|
    if row[0] == row2[0] && row[1] == row2[1] && row[3] == row2[3] && row[4] == row2[4] && row[5] == row2[5] && row[7] == row2[7]
      a = 'n'
      File.open("dup.txt", 'a+') { |file| file.write("#{row[0]},#{row[1]}; #{row[3]}; #{row[3]}; #{row[4]}; #{row[5]}; #{row[7]};") }
      File.open("dup.txt", 'a+') { |file| file.write("\n") }
      break
    end
  end
  if a=='y'
    File.open("loan.txt", 'a+') { |file| file.write("++GROUP") }
    File.open("loan.txt", 'a+') { |file| file.write("\n") }
       if row[1].length == 10
          File.open("loan.txt", 'a+') { |file| file.write("#{row[0]},#{row[1]};AMOUNT=#{row[3]}; ASOF=#{row[7]};#{row[4]}=#{row[5]};") }
       end
       if row[1].length == 9
          File.open("loan.txt", 'a+') { |file| file.write("#{row[0]},#{row[1]} ;AMOUNT=#{row[3]}; ASOF=#{row[7]};#{row[4]}=#{row[5]};")  }
       end
       File.open("loan.txt", 'a+') { |file| file.write("\n") }
       File.open("loan.txt", 'a+') { |file| file.write("ACTION=PROCESS;SUPPCHK=Y;DEPTDESK=POL2 ;. ")}
       File.open("loan.txt", 'a+') { |file| file.write("\n") }
       File.open("loan.txt", 'a+') { |file| file.write("++GROUPEND")}
       File.open("loan.txt", 'a+') { |file| file.write("\n") }
          if row[1].length == 10
            File.open("loan.txt", 'a+') { |file| file.write("#{row[0]},#{row[1]};FORCE=Y;.") }
          end
          if row[1].length == 9
            File.open("loan.txt", 'a+') { |file| file.write("#{row[0]},#{row[1]} ;FORCE=Y;.") }
          end
          File.open("loan.txt", 'a+') { |file| file.write("\n") }
          File.open("loan.txt", 'a+') { |file| file.write("QUIT") }
          File.open("loan.txt", 'a+') { |file| file.write("\n") }
  end
 end

