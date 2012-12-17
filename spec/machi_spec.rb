#encoding: cp932
$KODE="s"

require 'rubygems'
require './spec/test_helper'
require './lib/machi'

describe Machi do
	subject { Machi.new(debug: true) }

	describe "initialize section" do
		it "should have collect sheet amount" do
			subject.data_sheet.size.should == 1
		end
		it "should have collect sheet name" do
			subject.data_sheet[0].should =~ /斎藤/
		end
	end

	describe "reading section" do
		it "should have valuable 'data_container'" do
			subject.data_container.should_not be_nil
		end
		describe "should have fundamental" do
			before { subject.reader}
			it "have patient's data" do
				subject.fundamental[0][:code].should == "01234567"
				subject.fundamental[0][:shoshin].should == 1
				subject.fundamental[1][:code].should == "01122334"
			end
			it "data size is 2" do
				subject.fundamental.size.should == 2
			end
		end
	end
end