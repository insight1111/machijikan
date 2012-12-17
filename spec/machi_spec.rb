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
	  it "should have fundamental patient's data" do
	  	subject.reader
	  	subject.fundamental[0][:code].should == "01234567"
	  	subject.fundamental[0][:shoshin].should == 1
	  	subject.fundamental[1][:code].should == "01122334"
	  end
	end
end