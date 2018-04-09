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
    def inst(mock)
      @inst ||= Class.new do
        include AssOle::Rubify::Support::SendToOle
        attr_reader :ole
        def initialize(ole)
          @ole = ole
        end
      end.new(mock)
    end

    it 'send message to ole' do
      ole_mock = mock
      ole_mock.expects(:MethodName).with(:a1, :a2, :a3).returns(:result)
      inst(ole_mock).MethodName(:a1, :a2, :a3).must_equal :result
    end
  end

  describe AssOle::Rubify::Support::XmlType do
    module XmlTypeTest
      module SharedTests
        extend Minitest::Spec::DSL

        it '.get returns Nubmer' do
          AssOle::Rubify::Support::XmlType.get(1, ole_runtime_get).must_equal 'decimal'
        end

        it '.get returns nil' do
          AssOle::Rubify::Support::XmlType.get(:symbol, ole_runtime_get).must_be_nil
        end
      end
    end

    describe 'tests in serever context' do
      like_ole_runtime Runtimes::Ext
      include XmlTypeTest::SharedTests
    end

    describe 'tests in client context' do
      like_ole_runtime Runtimes::Thin
      include XmlTypeTest::SharedTests
    end
  end

  describe AssOle::Rubify::GenericWrapper do
    like_ole_runtime Runtimes::Ext
    include AssOle::Snippets::Shared::Array

    def klass
      AssOle::Rubify::GenericWrapper
    end

    def valid_ole_obj
      array
    end

    def ruby_ole_obj
      array(:hash).Get(0)
    end

    it 'include? Support::SendToOle' do
      assert klass.include? AssOle::Rubify::Support::SendToOle
    end

    it 'include? Support::XmlType' do
      assert klass.include? AssOle::Rubify::Support::XmlType
    end

    describe '#initialize' do
      it 'yelds self' do
        yielded = false
        klass.new(valid_ole_obj, ole_runtime_get) do |inst|
          inst.must_be_instance_of AssOle::Rubify::GenericWrapper
          yielded = true
        end
        yielded.must_equal true
      end

      describe 'fail' do
        it "if ole isn't WIN32OLE" do
          e = proc {
            klass.new(:invalid, ole_runtime_get)
          }.must_raise ArgumentError
          e.message.must_match %r{ole must be `WIN32OLE` instance}i
        end

        it "if ole isn't spawned ole_runtime" do
          e = proc {
            klass.new(ruby_ole_obj, ole_runtime_get)
          }.must_raise ArgumentError
          e.message.must_match %r{ole must be spawned by ole_runtime}i
        end
      end
    end

    describe 'tests with stubed klass' do
      def klass
        @klass ||= Class.new(super) do
          def initialize(ole, ole_runtime)
            @ole = ole
            @ole_runtime = ole_runtime
          end
        end
      end

      it '#xml_type' do
        inst = klass.new(:ole, :ole_runtime)
        AssOle::Rubify::Support::XmlType.expects(:get).with(:ole, :ole_runtime).returns(:xml_type)
        inst.xml_type.must_equal :xml_type
      end
    end
  end
end
