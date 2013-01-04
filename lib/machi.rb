#encoding: cp932
$KCODE="s"

require 'win32ole'
require 'pp'
require 'debugger'
require 'dbi'
require 'kconv'

class Machi
  attr_accessor :data_sheet, :data_container, :sh, :sheet_name, :db
  def initialize(options= {debug: false})
    @options = options
    @data_sheet=[]
    path="sheets"
    path+="_test" if options[:debug]
    Dir.glob("#{path}/*.xls").each do |sheet|
      @data_sheet << sheet
    end
    @data_container=[]
    @fundamental=[]
    @machijikan_kiso_data=[]
    @ex=nil
    @book=nil
    @sh=nil
    @db=DBI.connect("DBI:ODBC:machijikan",'admin','')
  end

  def reader
    begin
      @data_sheet.each do |sheet|
        @ex=WIN32OLE.new("Excel.Application")
        @sh=get_sheet_object(sheet)
        (2..get_last_line(@sh)).each do |line|
          _data_container = {}
          _data_container[:fundamental] = {
            code: (('00000000'+(@sh.cells(line,1)).value.to_i.to_s))[-8..-1], 
            drcode:          @sh.cells(line,2).value.to_i,
            shoshin:         @sh.cells(line,3).value.to_i,
            shokaijo:        @sh.cells(line,4).value.to_i,
            shinryoka:       @sh.cells(line,5).value.to_i,
            shinryoka_com:   @sh.cells(line,6).value.to_s,
            mokuteki:        @sh.cells(line,7).value.to_i,
            address:         @sh.cells(line,8).value.to_i
           }
          _data_container[:machijikan_kiso_data] = column_reader(@sh,line)
          _data_container[:machijikan] = calc_machijikan(_data_container[:machijikan_kiso_data],_data_container[:fundamental][:shinryoka])
          @data_container << _data_container
        end
      end
    rescue => e
      print e.message
    ensure
      @ex.quit
    end
  end

  def output
    begin
      @data_container.each do |d|
        f = d[:fundamental]
        db.do("insert into patients values(?,?,?,?,?,?,?,?);",
          f[:code], f[:drcode], f[:shoshin], f[:shokaijo],
          f[:shinryoka], f[:shinryoka_com], f[:mokuteki], f[:address])
        f = d[:machijikan_kiso_data]
        f.each do |_f|
          db.do("insert into machijikan_kisodata (code,type,uketsuke,start,end,min_time) values(?,?,?,?,?,?);",
            _f[0],_f[1],_f[2],_f[3],_f[4],_f[5])
        end
        f = d[:machijikan]
        f.each do |_f|
          # puts "insert into machijikan (code,shinryoka,type,timevalue) values(#{_f[0]},#{_f[1]},#{_f[2]},#{_f[3]});" 
          db.do("insert into machijikan (code,shinryoka,type,timevalue) values(?,?,?,?);",
            _f[0],_f[1],_f[2],_f[3])
        end
      end
    rescue DBI::DatabaseError => e      
      puts "An error occured."
      puts "Error code: #{e.err}".tosjis
      puts "Error message: #{e.message}".tosjis
    ensure
      db.disconnect if db
    end
  end

  private
  
    def getAbsolutePath filename
      fso = WIN32OLE.new('Scripting.FileSystemObject')
      return fso.GetAbsolutePathName(filename)
    end

    def get_sheet_object(filename)
      path=getAbsolutePath(filename)
      sh=@ex.workbooks.open(path).sheets("data")
    end

    def get_last_line(sheetobject)
      sheetobject.range("A1").end(-4121).row
    end

    def column_reader(sheet,line)
    	col=9
    	return_data=[]
    	until sheet.cells(line,col).value == nil && sheet.cells(line,col+1).value==nil && sheet.cells(line,col+2).value==nil
        # debugger
    		temp_data=[sheet.cells(line,1).value.to_i]
  			(col..col+3).each do |c|
          if c % 4 == 1
            temp_data << sheet.cells(line,c).value.to_i
            next
          end
          if sheet.cells(line,c).value
  				temp_data << convert_time(sheet.cells(line,c).value)
          else
          temp_data << nil
          end
  			end
        temp_data << [temp_data[2],temp_data[3],temp_data[4]].compact.min
  			return_data << temp_data  			
    		col+=4
    	end
      return_data.sort{|a,b| a[5] <=> b[5]}
    end

    def convert_time(time_string)
      # p time_string
      Time.local(2012,12,12,time_string[0..1].to_i, time_string[2..3].to_i)
    end
    
    def calc_machijikan(machijikan, shinryoka)
      return_data = []
      sogo_uketsuke          = nil
      shinryoka_uketsuke     = nil
      naishikyo_uketsuke     = nil
      naishikyo_monshin_end  = nil
      naishikyo_shochi_end   = nil
      gazo_uketsuke          = nil
      gazo_shochi_end        = nil
      shiharai_end           = nil
      machijikan.each_with_index do |m,i|
        # debugger if @options[:debug]
        type   = m[1]
        value = 0
        temp_shinryoka = nil
        begin
        case type
        when 1
          value = m[4]-m[2]
          sogo_uketsuke = m[2]
        when 2, 21..28
          shinryoka_uketsuke = m[2]
          next
        when 3 , 31..38
          value = m[3]-shinryoka_uketsuke
        when 4 , 41..48
          value = m[3]-machijikan[i-1][4]
        when 5 , 51..58
          value = m[3]-machijikan[i-1][4]
        when 6 , 61..62
          value = m[3]-m[2]
        when 7
          naishikyo_uketsuke = m[2]
          next
        when 71
          value = m[3] - naishikyo_uketsuke
          naishikyo_monshin_end = m[4]
        when 72
          value = m[3] - naishikyo_monshin_end
          naishikyo_shochi_end = m[4]
        when 73
          value = m[3] - naishikyo_shochi_end
        when 8
          gazo_uketsuke = m[2]
          next
        when 81, 810 .. 814
          value = m[3] - gazo_uketsuke
          gazo_shochi_end = m[4]
        when 82, 820 .. 824
          value = m[3] - gazo_shochi_end
        when 9
          value = m[3] - m[2]
        when 10
          value = m[4] - m[2]
        when 11
          value = m[3] - m[2]
        when 12
          value = m[4] - m[2]
          shiharai_end = m[4]
        end

        if type.between?(31,58)
          temp_shinryoka = get_shinryoka(type)
        end

        code = m[0]
        
        if type.between?(31,58)
          type = type.to_s[0].to_i
        end

        # puts "#{i}:#{value.to_s}"
        value = (value / 60).to_i if value 
        return_data << [code, (temp_shinryoka || shinryoka), type, value]
      rescue
        next
      end
      end
      return_data << [machijikan[0][0], nil, 99, ((shiharai_end - sogo_uketsuke) / 60).to_i] if shiharai_end && sogo_uketsuke  
      return_data
    end

    def get_shinryoka(type)
      case 
      when type > 800 
        type.to_s[2].to_i
      else
        type.to_s[1].to_i
      end
    end
end