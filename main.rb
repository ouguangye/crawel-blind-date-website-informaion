require "http"
require 'json'
require "spreadsheet/excel"

Spreadsheet.client_encoding="utf-8"
book=Spreadsheet::Workbook.new

class MyRequest
    def initialize(startage, endage, gender, startheight,endheight,salary)
        @startage = startage
        @endage = endage
        @gender = gender
        @startheight = startheight
        @endheight = endheight
        @salary = salary
    end

    def getData
        info = Array.new
        (1..50).each do |page|
            response = HTTP.get("https://m.7799520.com/api/recommend/wap/search/list/search",
                :params=>{
                    "startage"=>@startage,"endage"=>@endage, 
                    "gender"=>@gender, 
                    "startheight"=>@startheight,"endheight"=>@endheight,
                    "salary"=>@salary, 
                    "page"=>page
            })

            result = JSON.parse(response.body)

            break if result['error_code'] == -1 or result['data']['num'] == 0

            info = info + result['data']['list']
        end
        return info
    end

end


puts "\033[31m
.----------------.  .----------------.  .----------------.  .----------------.
| .--------------. || .--------------. || .--------------. || .--------------. |
| |   _____      | || |     ____     | || | ____   ____  | || |  _________   | |
| |  |_   _|     | || |   .'    `.   | || ||_  _| |_  _| | || | |_   ___  |  | |
| |    | |       | || |  /  .--.  \\  | || |  \\ \\   / /   | || |   | |_  \\_|  | |
| |    | |   _   | || |  | |    | |  | || |   \\ \\ / /    | || |   |  _|  _   | |
| |   _| |__/ |  | || |  \\ `-- '  /  | || |    \\ ' /     | || |  _| |___/ |  | |
| |  |________|  | || |   `.____.'   | || |     \\_/      | || | |_________|  | |
| |              | || |              | || |              | || |              | |
| '--------------' || '--------------' || '--------------' || '--------------' |
  '----------------'  '----------------'  '----------------'  '----------------'
\033[0m
"

puts "欢迎来到ruby相亲小程序, 下面来选择你理想型对象的特征，如果对相关条件没有具体要求，那么就直接按回车键\n "

print "请输入对象的理想年龄:  "
age = gets.chomp
user_startage = nil

user_endage = nil
if age != ''
    age = age.to_i
    user_startage = (age/10)*10
    user_endage = (age/10 + 1)*10
end

# p user_startage, user_endage

print "请输入对象的理想性别 (男性输入0, 女性输入1): "
user_gender=gets.chomp
if user_gender == ''
    user_gender = nil
else 
    user_gender = user_gender.to_i + 1
end

# p user_gender

print "请输入对象的理想身高: "
height = gets.chomp
user_startheight = nil
user_endheight = nil

if height != ''
    height = height.to_i
    user_startheight = (height/10)*10
    user_endheight = (height/10 + 1)*10
end

# p user_startheight, user_endheight

print "请输入对象的理想薪资(小于1万填0, 多于1万填1) : "
user_salary=gets.chomp

if user_salary == ''
    user_salary = nil
else
    user_salary = user_salary.to_i == 0 ? 3 : 4
end

# p user_salary

puts "即将根据您的要求筛选出用户信息，稍等片刻："
my_request = MyRequest.new(user_startage,user_endage,user_gender,user_startheight,user_endheight,user_salary)

data = my_request.getData
# puts data
# puts "展示完毕，一共#{data.length}条信息。"

if data.length == 0
    puts "抱歉, 这里没有找到匹配的信息"
else
    print "我们请输入表格的名称: "
    excelName = gets.chomp
    excelName = '你理想型对象信息汇总' if excelName == ''

    sheet1=book.create_worksheet :name => excelName
    sheet1.row(0)[0]="用户名"
    sheet1.row(0)[1]="性别"
    sheet1.row(0)[2]="头像"
    sheet1.row(0)[3]="年龄"
    sheet1.row(0)[4]="身高"
    sheet1.row(0)[5]="星座"
    sheet1.row(0)[6]="居住城市"
    sheet1.row(0)[7]="工资情况"
    sheet1.row(0)[8]="座右铭"

    index = 1
    gender_hash = {'1'=>'男', '2'=>'女'}
    time = Time.new

    data.each do |d|
        sheet1.row(index)[0] = d["username"]
        sheet1.row(index)[1] = gender_hash[d["gender"]]
        sheet1.row(index)[2] = d["avatar"]
        sheet1.row(index)[3] = time.year.to_i - d["birthdayyear"].to_i
        sheet1.row(index)[4] = d["height"]
        sheet1.row(index)[5] = d["astro"]
        sheet1.row(index)[6] = d["city"] == d["province"] ? d["city"] : d["province"] + d["city"]
        sheet1.row(index)[7] = d["salary"]
        sheet1.row(index)[8] = d["monolog"]
        index+=1
    end
    book.write "./#{excelName}.xls"
end