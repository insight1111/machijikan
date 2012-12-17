require 'rubygems'
require './spec/test_helper'
require './lib/hoge'

describe Hoge do
	subject { Hoge.new }

	it { subject.h.should == 'hoge'}
	it { subject.h2.should == 'hogehoge'}
end