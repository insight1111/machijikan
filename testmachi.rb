require './lib/machi'

m=Machi.new(debug: true)
m.reader
# name=m.data_sheet[0]
# sh=m.get_sheet_object(name)
p m.sheet_name