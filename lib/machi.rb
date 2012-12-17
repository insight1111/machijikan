#encoding: cp932
$KCODE="s"

require 'win32ole'
class Machi
	attr_accessor :data_sheet, :data_container, :fundamental, :sh, :sheet_name
	def initialize(options= {debug: false})
		@data_sheet=[]
		path="sheets"
		path+="_test" if options[:debug]
		Dir.glob("#{path}/*.xls").each do |sheet|
			@data_sheet << sheet
		end
		@data_container=[]
		@fundamental=[]
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
				end
				# p @fundamental
				@sheet_name=@sh.name
			end
		rescue => e
			print e
		ensure
			@ex.quit
		end
	end

	def getAbsolutePath filename
    fso = WIN32OLE.new('Scripting.FileSystemObject')
    return fso.GetAbsolutePathName(filename)
  end

  def get_sheet_object(filename)
		path=getAbsolutePath(filename)
		book=@ex.workbooks.open(path)
		sh=book.sheets("data")
  end

  def get_last_line(sheetobject)
  	sheetobject.range("A1").end(-4121).row
  end
end