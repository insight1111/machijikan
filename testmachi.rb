require './lib/machi'
require 'pp'

m=Machi.new(debug: true)
m.reader
pp m.data_container
# name=m.data_sheet[0]
# sh=m.get_sheet_object(name)
