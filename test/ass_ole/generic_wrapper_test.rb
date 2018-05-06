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
            klass.any_instance.expects(:mixins_blend)
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

          def mixins_blend; end

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

  describe AssOle::Rubify::GenericWrapper::Mixins do
    like_ole_runtime Runtimes::Ext

    describe 'Collection' do
      def inst(*args, **opts, &block)
        @inst ||= AssOle::Rubify::GenericWrapper
          .new(ole_coll(*args, **opts, &block), ole_runtime_get)
      end

      def ole_coll(*args, **opts, &block)
        fail 'Abstract method. Must return pure WIN32OLE collection'
      end

      describe 'OLE Array' do
        include AssOle::Snippets::Shared::Array

        def ole_coll(items = [1, 2, 3, 4, 5])
          @ole_coll ||= array items
        end

        it 'must include Indexable' do
          inst.singleton_class
            .include? AssOle::Rubify::GenericWrapper::Mixins::Collection::Indexable
        end

        describe '#each' do
          it 'without block' do
            inst.each.must_be_instance_of AssOle::Rubify::GenericWrapper
          end

          it 'with block' do
            times = 0

            inst.each do |item|
              times += 1
              item.must_equal times
            end.must_equal inst

            times.must_equal 5
          end
        end

        it '#each_with_index' do
          times = 0

          inst.each_with_index do |item, index|
            times += 1
            item.must_equal times
            index.must_equal item - 1
          end

          times.must_equal 5
        end

        it '#map' do
          inst.map(&:to_s).must_equal %w{1 2 3 4 5}
        end

        it '#size' do
          inst.size.must_equal 5
        end

        it '#count' do
          inst.count.must_equal 5
        end

        describe '#[index] when' do
          it 'array is empty returns nil' do
            inst([]).size.must_equal 0
            inst[-1].must_be_nil
            inst[0].must_be_nil
            inst[1].must_be_nil
          end

          it 'index out of the range returns nil' do
            inst.size.must_equal 5
            inst[5].must_be_nil
            inst[6].must_be_nil
          end

          it 'index in of the range returns item value' do
            inst.size.must_equal 5
            inst[0].must_equal 1
            inst[3].must_equal 4
            inst[4].must_equal 5
          end

          describe 'index < 0 will be reverse indexing' do
            it 'index in of the rage returns item value' do
              inst.size.must_equal 5
              inst[-1].must_equal 5
              inst[-2].must_equal 4
              inst[-3].must_equal 3
              inst[-4].must_equal 2
              inst[-5].must_equal 1
            end

            it 'index out of range returns nil' do
              inst.size.must_equal 5
              inst[-6].must_be_nil
            end
          end
        end
      end
    end
  end
end
