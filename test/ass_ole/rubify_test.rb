require "test_helper"
require 'ass_ole/snippets/shared'

module AssOle::RubifyTest
  describe 'Const' do
    it 'verify version number' do
      ::AssOle::Rubify::VERSION.wont_be_nil
    end
  end

  describe AssOle::Rubify::MdManagers do

  end

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

    it 'send message to ole' do
      ole_mock = mock
      ole_mock.expects(:MethodName).with(:a1, :a2, :a3).returns(:result)
      inst(ole_mock).MethodName(:a1, :a2, :a3).must_equal :result
    end

    it 'fail ArgumentError if method == ole' do
      e = proc {
        inst(nil).method_missing(:ole)
      }.must_raise ArgumentError
      e.message.must_match %r{All included `SendToOle` must respond_to\? `:ole`}i
    end

    describe '#method_missing' do
      it 'invocation sequnce' do
        invocation = sequence('invocation')
        ole_mock = mock

        inst.must_respond_to :_extract_args_
        inst.must_respond_to :_extract_opts_
        inst.must_respond_to :_fill_attributes_
        inst.must_respond_to :_wrapp_ole_result_
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

  describe AssOle::Rubify do
    like_ole_runtime Runtimes::Ext

    describe '.rubify' do
      it 'when ole is WIN32OLE' do
#FIXME        AssOle::Rubify::GenericWrapper.expects(:new)
#FIXME          .with(valid_ole_obj, ole_runtime_get).returns(:wrapper)
#FIXME        AssOle::Rubify.rubify(valid_ole_obj, ole_runtime_get).must_equal :wrapper
        skip "FIXME"
      end

      it 'when ole is StringInternal' do
        skip 'FIXME'
      end

      it 'when ole is Xml string' do
        skip 'FIXME'
      end

      it 'when ole is nil' do
        AssOle::Rubify.rubify(nil, ole_runtime_get).must_be_nil
      end

      it 'fail ArgumentError when ole is ivalid' do
        skip 'FIXME'
      end
    end

    it '#rubify' do
      AssOle::Rubify.expects(:rubify).with(:ole, ole_runtime_get).returns(:wrapper)
      Class.new do
        like_ole_runtime Runtimes::Ext
        include AssOle::Rubify
      end.new.rubify(:ole).must_equal :wrapper
    end
  end
end
