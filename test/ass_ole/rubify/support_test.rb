require 'test_helper'

module AssOle::RubifyTest
  describe AssOle::Rubify::Support::SendToOle do
    def inst(mock_ole = nil)
      @inst ||= Class.new do
        include AssOle::Rubify::Support::SendToOle
        attr_accessor :ole
        def initialize(ole)
          @ole = ole
        end
      end.new(mock_ole)
    end

    it 'fail ArgumentError if method == ole' do
      e = proc {
        inst(nil).method_missing(:ole)
      }.must_raise ArgumentError
      e.message.must_match %r{All included `SendToOle` must respond_to\? `:ole`}i
    end

    describe '#method_missing' do
      describe 'invocation sequence' do
        it 'if method reader' do
          invocation = sequence('invocation')
          ole_mock = mock

          inst.must_respond_to :_extract_args_
          inst.must_respond_to :_extract_opts_
          inst.must_respond_to :_fill_attributes_
          inst.must_respond_to :_wrapp_ole_result_
          inst.must_respond_to :_writer_missing_
          yelded = nil

          inst.expects(:_extract_args_).with([:a1,:a2, :a3])
            .in_sequence(invocation).returns([:a1, :a2, :a3])
          ole_mock.expects(:send).with(:Method, :a1, :a2, :a3)
             .in_sequence(invocation).returns(:result)
          inst.expects(:_extract_opts_).with({op1: 1, op2: 2, op3: 3})
            .in_sequence(invocation).returns(:opts)
          inst.expects(:_fill_attributes_).with(:result, :opts)
            .in_sequence(invocation).returns(:result)
          inst.expects(:_wrapp_ole_result_).with(:result)
            .in_sequence(invocation).returns(:result)
          inst.ole = ole_mock

          inst.Method(:a1, :a2, :a3, op1: 1, op2: 2, op3: 3) do |val|
            yelded = true
            val.must_equal :result
            nil
          end.must_equal :result
        end

        it 'if method writer' do
          inst(:ole).must_respond_to :_extract_args_
          inst.must_respond_to :_extract_opts_
          inst.must_respond_to :_fill_attributes_
          inst.must_respond_to :_wrapp_ole_result_
          inst.must_respond_to :_writer_missing_

          inst.expects(:_extract_args_).never
          inst.expects(:_extract_opts_).never
          inst.expects(:_fill_attributes_).never
          inst.expects(:_wrapp_ole_result_).never
          inst.expects(:_writer_missing_).with(:Method=, :value).returns(:value)
          yielded = false

          (inst.Method = :value).must_equal :value
          yielded.must_equal false
        end

        it '#_writer_missing_' do
          invocation = sequence('invocation')
          ole_mock = mock
          inst.expects(:_extract_ole_).with(:value)
            .in_sequence(invocation).returns(:extracted_value)
          ole_mock.expects(:send).with(:Method=, :extracted_value)
            .in_sequence(invocation).returns(:extracted_value)
          inst.ole = ole_mock
          inst._writer_missing_(:Method=, :value).must_equal :value
        end
      end
    end

    describe '#_writer_missing_' do
      describe 'always returns the same value was got' do
        it 'with single value' do
          ole = mock
          ole.expects(:send).with(:Method=, :value).returns(:othe_value)
          (inst(ole).Method = :value).must_equal :value
        end

        it 'with multiple values' do
          ole = mock
          ole.expects(:send).with(:Method=, [1,2,3]).returns(:othe_value)
          (inst(ole).Method = 1,2,3).must_equal [1,2,3]
        end
      end
    end

    it '#_extract_args_' do
      inst.respond_to? :_extract_ole_
      inst.expects(:_extract_ole_).with(1).returns(:'1')
      inst.expects(:_extract_ole_).with(2).returns(:'2')
      inst.expects(:_extract_ole_).with(3).returns(:'3')

      inst._extract_args_([1,2,3]).must_equal([:'1', :'2', :'3'])
    end

    it '#_extract_opts_' do
      inst.respond_to? :_extract_ole_
      inst.expects(:_extract_ole_).with(1).returns(:'1')
      inst.expects(:_extract_ole_).with(2).returns(:'2')
      inst.expects(:_extract_ole_).with(3).returns(:'3')

      inst._extract_opts_({o1: 1, o2: 2, o3: 3})
        .must_equal({o1: :'1', o2: :'2', o3: :'3'})
    end

    it '#_extract_ole_' do
      inst._extract_ole_('Abstract method').must_equal 'Abstract method'
    end

    it '#_wrapp_ole_result_' do
      inst._wrapp_ole_result_('Abstract method').must_equal 'Abstract method'
    end
  end

  describe AssOle::Rubify::Support::DuckTyping do
    it 'method_missing' do
      wrapper = mock
      wrapper.expects(:ole_respond_to?).with(:Foo).returns(true)
      wrapper.expects(:ole_respond_to?).with(:Bar).returns(false)
      inst = AssOle::Rubify::Support::DuckTyping.new(wrapper)
      inst.Foo?.must_equal true
      inst.Bar?.must_equal false
    end
  end
end
