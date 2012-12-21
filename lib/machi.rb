#encoding: cp932
$KCODE="s"

require 'win32ole'
require 'pp'
# require 'ruby-debug'
class Machi
  attr_accessor :data_sheet, :data_container, :fundamental, :sh, :sheet_name, :machijikan_data
  def initialize(options= {debug: false})
    @data_sheet=[]
    path="sheets"
    path+="_test" if options[:debug]
    Dir.glob("#{path}/*.xls").each do |sheet|
      @data_sheet << sheet
    end
    @data_container=[]
    @fundamental=[]
    @machijikan_data=[]
    @ex=nil
    @book=nil
    @sh=nil
  end

  def reader
    begin
      @data_sheet.each do |sheet|
        @ex=WIN32OLE.new("Excel.Application")
        @sh=get_sheet_object(sheet)
        (2..get_last_line(@sh)).each do |line|
          @fundamental << {
            code: (('00000000'+(@sh.cells(line,1)).value.to_i.to_s))[-8..-1], 
            drcode:          @sh.cells(line,2).value.to_i,
            shoshin:         @sh.cells(line,3).value.to_i,
            shokaijo:        @sh.cells(line,4).value.to_i,
            shinryoka:       @sh.cells(line,5).value.to_i,
            shinryoka_com:   @sh.cells(line,6).value.to_s,
            mokuteki:        @sh.cells(line,7).value.to_i,
            address:         @sh.cells(line,8).value.to_i
           }
          @machijikan_data << column_reader(@sh,line)
        end
      end
    rescue => e
      print e.message
    ensure
      @ex.quit
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
      return_data
    end

    def convert_time(time_string)
      # p time_string
      Time.local(2012,12,12,time_string[0..1].to_i, time_string[2..3].to_i)
    end
    
end