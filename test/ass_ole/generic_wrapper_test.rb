require 'test_helper'
require 'ass_ole/snippets/shared'
module AssOle::RubifyTest
  describe AssOle::Rubify::GenericWrapper do
    like_ole_runtime Runtimes::Ext
    include AssOle::Snippets::Shared::Array

    def other_ole_runtime
      Class.new do
        like_ole_runtime Runtimes::Thin
        include AssOle::Snippets::Shared::Array
      end.new
    end

    def klass
      AssOle::Rubify::GenericWrapper
    end

    def valid_ole_obj
      array
    end

    def invalid_ole_obj
      other_ole_runtime.array
    end

    it '#to_s' do
      klass.new(valid_ole_obj, ole_runtime_get).to_s.must_match %r{Массив|Array}i
    end

    it '#quack' do
      inst = klass.new(valid_ole_obj, ole_runtime_get)
      inst.quack.Get?.must_equal true
      inst.quack.Fake?.must_equal false
    end

    it '#glob_context' do
      klass.new(valid_ole_obj, ole_runtime_get).glob_context
        .must_be_instance_of AssOle::Rubify::GlobContex
    end

    it 'include? Support::SendToOle' do
      assert klass.include? AssOle::Rubify::Support::SendToOle
    end

    describe '#_wrapp_ole_result_' do
      it 'returns GenericWrapper' do
        inst = klass.new(valid_ole_obj, ole_runtime_get)
        actual = inst._wrapp_ole_result_(valid_ole_obj)
        actual.must_be_instance_of AssOle::Rubify::GenericWrapper
        actual.to_s.must_match %r{Array|Массив}
        actual.ole_runtime.must_equal ole_runtime_get
      end

      it 'returns value' do
        inst = klass.new(valid_ole_obj, ole_runtime_get)
        inst._wrapp_ole_result_(invalid_ole_obj).must_be_instance_of WIN32OLE
        inst._wrapp_ole_result_(nil).must_be_nil
        inst._wrapp_ole_result_('str').must_equal 'str'
      end
    end

    describe '#_extract_ole_' do
      it 'when value is GenericWrapper' do
        inst = klass.new(valid_ole_obj, ole_runtime_get)
        actual = inst._extract_ole_(inst)
        actual.must_be_instance_of WIN32OLE
        sTring(actual).must_match %r{Массив|Array}
      end

      it "when value isn't GenericWrapper" do
        inst = klass.new(valid_ole_obj, ole_runtime_get)
        inst._extract_ole_(:not_wrapper).must_equal :not_wrapper
      end
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
            klass.new(invalid_ole_obj, ole_runtime_get)
          }.must_raise ArgumentError
          e.message.must_match %r{ole must be spawned by ole_runtime}i
        end
      end
    end

    describe 'tests with stubed klass' do
      def klass
        @klass ||= Class.new(super) do
          def initialize(ole, ole_runtime)
            super
          end

          def verify!; end
        end
      end

      it '#xml_type' do
        ole_runtime = mock
        ole_runtime.responds_like(Runtimes::Ext)
        ole_runtime.expects(:xml_type_get).with(:ole).returns(:xml_type)
        inst = klass.new(:ole, ole_runtime)
        inst.xml_type.must_equal :xml_type
      end

      it '#to_string_internal' do
        ole_runtime = mock
        ole_runtime.responds_like(Runtimes::Ext)
        ole_runtime.expects(:to_string_internal).with(:ole).returns(:str_internal)
        inst = klass.new(:ole, ole_runtime)
        inst.to_string_internal.must_equal :str_internal
      end
    end
  end
end
