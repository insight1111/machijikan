#encoding: cp932
$KODE="s"

require 'rubygems'
require './spec/test_helper'
require './lib/machi'

# data structures
# 
# data_container[Array]
#   fundamental:[Hash]  ... personal data[Array]
#   machijikan_kiso_data:[Hash] ... original data[Array]
#   machijikan:[Hash] ... culculate machijikan data[Array]

describe Machi do
  subject { Machi.new(debug: true) }

  describe "initialize section" do
    it "should have collect sheet amount" do
      subject.data_sheet.size.should == 1
    end
    it "should have collect sheet name" do
      subject.data_sheet[0].should =~ /test/
    end
  end

  describe "reading section" do
    it "should have valuable 'data_container'" do
      subject.data_container.should_not be_nil
    end
    describe "should have fundamental" do
      before { subject.reader}
      it "collect patient's data" do
        subject.fundamental[0][:code].should == "01234567"
        subject.fundamental[0][:shoshin].should == 1
        subject.fundamental[1][:code].should == "01122334"
      end
      it "data size is 2" do
        subject.fundamental.size.should == 2
      end
    end

    describe "column reader" do
      before { subject.reader}
      it "should have three data" do
        subject.machijikan_kiso_data[0].size.should == 3
        subject.machijikan_kiso_data[1].size.should == 1
      end
      it "should firstData is integer" do
        subject.machijikan_kiso_data[0][0][0].should be_kind_of(Integer)
      end
      it "should third data is time format" do
      	subject.machijikan_kiso_data[0][0][2].should be_kind_of(Time)
      end
      # data structures
      #   code
      #   type(koumoku)
      #   uketsuke
      #   start
      #   end
      #   min_time...which is most early time??
      it "a piece of data should have six data" do
        subject.machijikan_kiso_data[0][0].size.should == 6
      end
      it "min_time is 9:25" do
        subject.machijikan_kiso_data[0][0][5].should == Time.local(2012,12,12,9,25)
      end
    end
  end

  # describe "output database" do
  #   before do
  #     subject.reader
  #     subject.output_database
  #   end
  #   it "should have dataconnection" do
  #     subject.database_connection.should_not be_nil
  #   end
  # end
end