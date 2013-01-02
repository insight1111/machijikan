#encoding: utf-8
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
        subject.data_container[0][:fundamental][:code].should == "01234567"
        subject.data_container[0][:fundamental][:shoshin].should == 1
        subject.data_container[1][:fundamental][:code].should == "01122334"
      end
      it "data size is 2" do
        subject.data_container.size.should == 2
      end
    end

    describe "column reader" do
      before { subject.reader}
      it "should have three data" do
        subject.data_container[0][:machijikan_kiso_data].size.should == 3
        subject.data_container[1][:machijikan_kiso_data].size.should == 1
      end
      it "should firstData is integer" do
        subject.data_container[0][:machijikan_kiso_data][0][0].should be_kind_of(Integer)
      end
      it "should third data is time format" do
      	subject.data_container[0][:machijikan_kiso_data][0][2].should be_kind_of(Time)
      end
      # machijikan_kiso_data structures
      #   code
      #   type(koumoku)
      #   uketsuke
      #   start
      #   end
      #   min_time...which is most early time??
      it "a piece of data should have six data" do
        subject.data_container[0][:machijikan_kiso_data][0].size.should == 6
      end
      it "min_time is 9:25" do
        subject.data_container[0][:machijikan_kiso_data][0][5].should == Time.local(2012,12,12,9,25)
      end

      #machijikan details
      # 1..終わり-受付
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