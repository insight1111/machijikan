#encoding: s
$KCODE="s"
# require 'pp'
require 'dbi'
require 'date'
require 'win32ole'
require 'ruby-debug'

def msgbox(msg,title)
	wsh=WIN32OLE.new("WScript.shell")
	wsh.popup(msg,0,title,0+64)
end


def excelopen
begin
ex    =WIN32OLE.new("Excel.Application")
userprofile=ENV['userprofile'].gsub('\\','/')
p path = userprofile+"/デスクトップ/test日報連絡\表\.xls"
# debugger:
book  =ex.workbooks.open(path)
p book.name
# book  =ex.workbooks.open('z:/test日報連絡表.xls')
sh    =book.sheets("減免")
shkin =book.sheets("無菌")
yield sh,shkin 
book.save
ensure
	book.close
end
end

begin
	excelopen do |sh,shkin|
#部屋情報を読み取り
dba=DBI.connect("DBI:ODBC:genmen",'admin','')
r=dba.execute("select * from rooms")
roomdata={}
r.fetch do |rr|
	roomdata[rr["physical_name"]]={'logic_name'=>rr["logic_name"],'fee'=>rr["fee"]}
end
r.finish

ka={1=>'内科',3=>"外科",4=>"整形外科",5=>"泌尿器科",7=>"産婦人科",12=>"放射線",10=>"歯",
  16=>"血液内科",17=>"消化器科",18=>"循環器科",19=>"脳神経外科",
  22=>"腎臓内科",41=>"総合診療科",42=>"脳神経内科",51=>"呼吸器科",
  71=>"糖尿病科",81=>"睡眠呼吸障害センター"}


# db=DBI.connect("DBI:ODBC:srv13",'viewer','')

data=[]


# 最終行を取得

line    = 3
kinline =10
while sh.cells(line,3).value!=nil 
	line += 1
end
while shkin.cells(kinline,3).value!=nil 
	kinline += 1
end

p sh.name
p shkin.name
p line
p kinline

#accessから取得
sql=<<EOF
select * from
reduce_data
where
putted=0
EOF
r        =dba.execute(sql)
tempdata ={}
data     =[]
flag     ={}
flag["jusho"]  =false
flag["mukin"]  =false
flag["genmen"] =false
r.fetch do |row|
	tempdata["id"]           = row["id"]
	tempdata["beforeroomno"] = row["beforeroomno"]
	tempdata["roomno"]       = row["roomno"]
	tempdata["shinsei_date"] = row["shinsei_date"].strftime("%m/%d")
	tempdata["start_day"]    = row["start_day"].strftime("%m/%d")
	tempdata["code"]         = row["code"]
	tempdata["name"]         = row["name"]
	tempdata["kana"]         = row["kana"]
	if row["end_day"]
	tempdata["end_day"]=row["end_day"].strftime("%m/%d")
	else
	tempdata["end_day"]=""
	end
	tempdata["nyuin_date"]=row["nyuin_date"].strftime("%m/%d") if row["nyuin_date"]
	if row["taiin_date"]
	tempdata["taiin_date"]=row["taiin_date"].strftime("%m/%d")
	end
	tempdata["reduce_fee"]   =row["reduce_fee"]
	tempdata["reduce_id"]    =row["reduce_id"]
	case tempdata["reduce_id"]
	when 1..2
		flag["genmen"]=true
	when 5..8
		flag["mukin"] =true
	when 9..10
		flag["jusho"] =true
	end
	tempdata["shinsei_name"]   = row["shinsei_name"]
	tempdata["reduce_fee"]     = row["reduce_fee"]
	tempdata["reporter_id"]    = row["reporter_id"]
	tempdata["reason_id"]      = row["reason_id"]
	tempdata["reason_comment"] = row["reason_comment"]
	data << tempdata
	tempdata={}
end
r.finish

data.each do |d|
	if [1, 2, 4, 9, 10].index(d["reduce_id"].to_i)
		# debugger
		room    =d["roomno"]
		roomfee =roomdata[room]['fee']
		roomlog =roomdata[room]['logic_name']
		fee=case d["reduce_id"]
				when 1
					"\\#{roomfee}→\\0"
				when 2
					"\\#{roomfee}→\\#{roomfee-d['reduce_fee']}"
				when 4
					"\\0→\\#{roomfee}"
				when 9
					"重症加算あり"
				when 10
					"重症加算なし"
				end
		sh.cells(line,1).value=d["shinsei_date"]
		sh.cells(line,2).value=d["code"]
		sei=d["name"].split(" ")[0]
		mei=d["name"].split(" ")[1..-1].join.gsub(/( |　)/,'')
		sh.cells(line,3).value=sei+" "+mei
		sh.cells(line,4).value=d["shinsei_name"]
		sh.cells(line,5).value="'(#{room[0..3]+'-'+room[-2..-1]})"
		sh.cells(line,6).value=fee
		if d['nyuin_date'] == d['start_day']
			nyuin_flag="ad"
		else
			nyuin_flag = ""
		end
		sh.cells(line,7).value="#{d['start_day']}#{nyuin_flag}～#{d['end_day']}"
		sql="select name from reporters where id=#{d['reporter_id']}"
		r=dba.execute(sql)
		sh.cells(line,8).value=r.fetch[0]
		r.finish
		sh.cells(line,11).value=d["reason_id"].to_s+"."+d["reason_comment"]
		sh.range(sh.cells(line,1),sh.cells(line,10)).borders.linestyle=true
		line=line+1
		sql="update reduce_data set putted =-1 where id = #{d['id'].to_i}"
		dba.do(sql)
	elsif [5,6,7,8].index(d["reduce_id"].to_i)
		# debugger
		room=d["roomno"]
		beforeroomno=d["beforeroomno"] ||= ""
		# roomlog=roomdata[room]['logic_name']
		fee=case d["reduce_id"]
				when 5
					"無菌なし"
				when 6
					"無菌開始"
				when 7
					"無菌続行"
				when 8
					"無菌中止"
				end
		shkin.cells(kinline,1).value=d["shinsei_date"]
		roomvalue=""
		if beforeroomno!=nil
			roomvalue="#{beforeroomno}→"
		end
		roomvalue+="#{room[0..3]+'-'+room[-2..-1]}"
		# shkin.cells(kinline,2).value= "'"+roomvalue
		shkin.cells(kinline,2).value= "'"+beforeroomno
		shkin.cells(kinline,3).value=d["code"]
		# debugger
		sei=d["name"].split(" ")[0]
		mei=d["name"].split(" ")[1..-1].join.gsub(/( |　)/,'')
		shkin.cells(kinline,4).value=sei+" "+mei
		shkin.cells(kinline,5).value=d["shinsei_name"].split(/( |　)/)[0]
		if d['nyuin_date'] == d['start_day']
			nyuin_flag="ad"
		else
			nyuin_flag = ""
		end
		datetemp=d['start_day'].split("/")
		datevalue=datetemp[0]+"月"+datetemp[1]+"日"
		shkin.cells(kinline,6).value="#{datevalue}#{nyuin_flag}"
		shkin.cells(kinline,7).value=fee
		sql="select name from reporters where id=#{d['reporter_id']}"
		r=dba.execute(sql)
		shkin.cells(kinline,8).value=r.fetch[0]
		r.finish
		shkin.range(shkin.cells(kinline,1),shkin.cells(kinline,9)).borders.linestyle=true
		kinline=kinline+1
		sql="update reduce_data set putted =-1 where id = #{d['id'].to_i}"
		dba.do(sql)
	end
end
# p line
dba.disconnect if dba
msgbox('出力完了しました','出力完了しました')
end
rescue => error
	puts $@
	puts error.backtrace
	puts error
	open('c:/genmen_error.txt','w'){ |f|
	f.puts $@
	f.puts error.backtrace
	f.puts error.message
	}
	puts "エラー内容をc:/genmen_error.txtに書き込みました"
end
