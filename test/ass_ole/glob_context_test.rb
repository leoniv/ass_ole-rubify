require "test_helper"

module AssOle::RubifyTest
  describe AssOle::Rubify::GlobContex do
    like_ole_runtime Runtimes::Ext

    def inst
      @inst = AssOle::Rubify::GlobContex.new(ole_runtime_get)
    end

    it '.superclass' do
      AssOle::Rubify::GlobContex.superclass
        .must_equal AssOle::Rubify::GenericWrapper
    end

    it '#newObject' do
      arr = inst.newObject('Array') do |a|
        3.times do |i|
          a.Add(i)
        end
      end

      arr.must_be_instance_of AssOle::Rubify::GenericWrapper
      arr.Count.must_equal 3

      3.times do |i|
        arr.Get(i).must_equal i
      end
    end
  end

  describe '.like_rubify_runtime' do
    describe 'in spec class' do
      like_rubify_runtime Runtimes::Ext

      it 'smoky with Array' do
        arr = newObject('Array') do |a|
          3.times do |i|
            a.Add(i)
          end
        end

        arr.must_be_instance_of AssOle::Rubify::GenericWrapper
        arr.Count.must_equal 3

        3.times do |i|
          arr.Get(i).must_equal i
        end
      end
    end

    describe Module do
      def _module
        @_module ||= Module.new do
          like_rubify_runtime Runtimes::Ext
        end
      end
      alias_method :inst, :_module

      it '#real_win_path' do
        inst.must_respond_to :real_win_path
        inst.real_win_path(__FILE__).must_match %r{ass_ole\\glob_context_test}
      end

      it '#argv' do
        inst.must_respond_to :argv
      end

      it 'instance has #glob_contex' do
        inst.glob_context.must_be_instance_of AssOle::Rubify::GlobContex
      end

      describe 'all transparecy sends to #glob_context' do
        it 'smoky' do
          arr = inst.newObject('Array') do |a|
            3.times do |i|
              a.Add(i)
            end
          end

          arr.must_be_instance_of AssOle::Rubify::GenericWrapper
          arr.Count.must_equal 3

          3.times do |i|
            arr.Get(i).must_equal i
          end
        end

        it 'mocked' do
          inst.glob_context.expects(:newObject).with('Array', {op1: 1, op2: 2})
          inst.newObject('Array', op1: 1, op2: 2)
        end
      end
    end

    describe Class do
      def klass
        @klass ||= Class.new do
          like_rubify_runtime Runtimes::Ext
        end
      end

      def inst
        @inst ||= klass.new
      end

      it '#real_win_path' do
        inst.must_respond_to :real_win_path
        inst.real_win_path(__FILE__).must_match %r{ass_ole\\glob_context_test}
      end

      it '#argv' do
        inst.must_respond_to :argv
      end

      it 'instance has #glob_contex' do
        inst.glob_context.must_be_instance_of AssOle::Rubify::GlobContex
      end

      describe 'all transparecy sends to #glob_context' do
        it 'smoky' do
          arr = inst.newObject('Array') do |a|
            3.times do |i|
              a.Add(i)
            end
          end

          arr.must_be_instance_of AssOle::Rubify::GenericWrapper
          arr.Count.must_equal 3

          3.times do |i|
            arr.Get(i).must_equal i
          end
        end

        it 'mocked' do
          inst.glob_context.expects(:newObject).with('Array', {op1: 1, op2: 2})
          inst.newObject('Array', op1: 1, op2: 2)
        end
      end
    end
  end

  describe AssOle::Rubify::GlobContex::Mixins::MixedContext do
    module MixedContextSharedTests
      extend Minitest::Spec::DSL

      it 'must include MixedContext' do
        glob_context.singleton_class
          .include?(AssOle::Rubify::GlobContex::Mixins::MixedContext)
          .must_equal true
      end

      it '#Array' do
        glob_context.Array(1,2,3,4).must_be_instance_of AssOle::Rubify::GenericWrapper
        glob_context.Array(1,2,3,4).to_s.must_match %r{Array|Массив}
      end

      it '#Map' do
        actual = glob_context.Map(:'key 1' => 1, :'key 2' => 2)
        actual.must_be_instance_of AssOle::Rubify::GenericWrapper
        actual.to_s.must_match %r{Map|Соответствие}
      end

      it '#Structure' do
        actual = glob_context.Structure(:key1 => 1, :key2 => 2)
        actual.must_be_instance_of AssOle::Rubify::GenericWrapper
        actual.to_s.must_match %r{Structure|Структура}
      end

      it '#Type' do
        actual = glob_context.Type('CatalogRef.Catalog1')
        actual.must_be_instance_of AssOle::Rubify::GenericWrapper
        actual.to_s.must_match %r{Catalog1}
        actual.xml_type.must_equal 'Type'
      end
    end

    module ClientContextSharedTests
      extend Minitest::Spec::DSL

      it 'must include ClientContext' do
        glob_context.singleton_class
          .include?(AssOle::Rubify::GlobContex::Mixins::ClientContext)
          .must_equal true
      end
    end

    module ServerContextSharedTests
      extend Minitest::Spec::DSL

      it 'must include ServerContext' do
        glob_context.singleton_class
          .include?(AssOle::Rubify::GlobContex::Mixins::ServerContext)
          .must_equal true
      end

      it '#Query' do
        actual = glob_context.Query('select 2')
        actual.must_be_instance_of AssOle::Rubify::GenericWrapper
        actual.to_s.must_match %r{Query|Запрос}
      end

      it '#ValueTabale' do
        actual = glob_context.ValueTable(:col1, :col2, :col3)
        actual.must_be_instance_of AssOle::Rubify::GenericWrapper
        actual.to_s.must_match %r{ValueTable|ТаблицаЗначений}
      end
    end

    describe 'External global context' do
      def glob_context
        AssOle::Rubify::GlobContex.new(Runtimes::Ext)
      end

      include MixedContextSharedTests
      include ServerContextSharedTests
    end

    describe 'Thin client global context' do
      def glob_context
        AssOle::Rubify::GlobContex.new(Runtimes::Thin)
      end

      include MixedContextSharedTests
      include ClientContextSharedTests
    end

    describe 'Thick client global context' do
      def glob_context
        AssOle::Rubify::GlobContex.new(Runtimes::Thick)
      end

      include MixedContextSharedTests
      include ClientContextSharedTests
      include ServerContextSharedTests
    end
  end
end
