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
p path = userprofile+"/�f�X�N�g�b�v/test����A��\�\\.xls"
# debugger:
book  =ex.workbooks.open(path)
p book.name
# book  =ex.workbooks.open('z:/test����A���\.xls')
sh    =book.sheets("����")
shkin =book.sheets("����")
yield sh,shkin 
book.save
ensure
	book.close
end
end

begin
	excelopen do |sh,shkin|
#��������ǂݎ��
dba=DBI.connect("DBI:ODBC:genmen",'admin','')
r=dba.execute("select * from rooms")
roomdata={}
r.fetch do |rr|
	roomdata[rr["physical_name"]]={'logic_name'=>rr["logic_name"],'fee'=>rr["fee"]}
end
r.finish

ka={1=>'����',3=>"�O��",4=>"���`�O��",5=>"��A���",7=>"�Y�w�l��",12=>"���ː�",10=>"��",
  16=>"���t����",17=>"�������",18=>"�z���",19=>"�]�_�o�O��",
  22=>"�t������",41=>"�����f�É�",42=>"�]�_�o����",51=>"�ċz���",
  71=>"���A�a��",81=>"�����ċz��Q�Z���^�["}


# db=DBI.connect("DBI:ODBC:srv13",'viewer','')

data=[]


# �ŏI�s���擾

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

#access����擾
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
					"\\#{roomfee}��\\0"
				when 2
					"\\#{roomfee}��\\#{roomfee-d['reduce_fee']}"
				when 4
					"\\0��\\#{roomfee}"
				when 9
					"�d�ǉ��Z����"
				when 10
					"�d�ǉ��Z�Ȃ�"
				end
		sh.cells(line,1).value=d["shinsei_date"]
		sh.cells(line,2).value=d["code"]
		sei=d["name"].split(" ")[0]
		mei=d["name"].split(" ")[1..-1].join.gsub(/( |�@)/,'')
		sh.cells(line,3).value=sei+" "+mei
		sh.cells(line,4).value=d["shinsei_name"]
		sh.cells(line,5).value="'(#{room[0..3]+'-'+room[-2..-1]})"
		sh.cells(line,6).value=fee
		if d['nyuin_date'] == d['start_day']
			nyuin_flag="ad"
		else
			nyuin_flag = ""
		end
		sh.cells(line,7).value="#{d['start_day']}#{nyuin_flag}�`#{d['end_day']}"
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
					"���ۂȂ�"
				when 6
					"���ۊJ�n"
				when 7
					"���ۑ��s"
				when 8
					"���ے��~"
				end
		shkin.cells(kinline,1).value=d["shinsei_date"]
		roomvalue=""
		if beforeroomno!=nil
			roomvalue="#{beforeroomno}��"
		end
		roomvalue+="#{room[0..3]+'-'+room[-2..-1]}"
		# shkin.cells(kinline,2).value= "'"+roomvalue
		shkin.cells(kinline,2).value= "'"+beforeroomno
		shkin.cells(kinline,3).value=d["code"]
		# debugger
		sei=d["name"].split(" ")[0]
		mei=d["name"].split(" ")[1..-1].join.gsub(/( |�@)/,'')
		shkin.cells(kinline,4).value=sei+" "+mei
		shkin.cells(kinline,5).value=d["shinsei_name"].split(/( |�@)/)[0]
		if d['nyuin_date'] == d['start_day']
			nyuin_flag="ad"
		else
			nyuin_flag = ""
		end
		datetemp=d['start_day'].split("/")
		datevalue=datetemp[0]+"��"+datetemp[1]+"��"
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
msgbox('�o�͊������܂���','�o�͊������܂���')
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
	puts "�G���[���e��c:/genmen_error.txt�ɏ������݂܂���"
end
