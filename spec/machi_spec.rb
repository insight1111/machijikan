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
      it "data size is 3" do
        subject.data_container.size.should == 3
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
      #   0:code
      #   1:type(koumoku)
      #   2:uketsuke
      #   3:start
      #   4:end
      #   5:min_time...which is most early time??
      it "a piece of data should have six data" do
        subject.data_container[0][:machijikan_kiso_data][0].size.should == 6
      end
      it "min_time is 9:25" do
        subject.data_container[0][:machijikan_kiso_data][0][5].should == Time.local(2012,12,12,9,25)
      end
      it "sorted by min_time" do
        subject.data_container[0][:machijikan_kiso_data][1][5].should == Time.local(2012,12,12,9,40)
      end

      # 待ち時間の定義
      #
      # 1...終わり-受付
      # 3...3開始-2受付
      # 4...4開始-直前終わり
      # 5...5開始-直前終わり
      # 6...6開始-6受付
      # 71..71開始-7受付
      # 72..72開始-71終了
      # 73..73開始-72終了
      # 81..81開始-8受付
      # 82..82開始-81終了
      # 9...9開始-9受付
      # 10..10終了-10受付
      # 11..11開始-11受付
      # 12..12終了-12受付
      # 99..滞在時間..12終了-1受付-12-13

      # 待ち時間データ構造定義
      # 
      # code(paitent id)
      # shinryoka(allow null)
      # machijikan_id(above definition value)
      # value(waiting time)

      it "have machijikan" do
        subject.data_container[0][:machijikan].should_not be_nil
      end
      it "a piece of machijikan_data is array and contains four data" do
        p subject.data_container[0][:machijikan][0]
        subject.data_container[0][:machijikan][0].should be_a_kind_of(Array)
        subject.data_container[0][:machijikan][0].size.should == 4
      end
      it "machijikan first data size is 3" do
        subject.data_container[0][:machijikan].size.should == 2
      end
      it "machijikan[0] first machijikan is 5" do
        subject.data_container[0][:machijikan][0][3].should == 5
      end
    end
  end

  describe "output database" do
    before do
      subject.reader
      subject.output
    end
    it "should have dataconnection" do
      subject.db.should_not be_nil
    end
  end
end