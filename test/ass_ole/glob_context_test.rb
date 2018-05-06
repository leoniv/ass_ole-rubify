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

      it '#type_get' do
        actual = glob_context.type_get('CatalogRef.Catalog1')
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

      it '#query_get' do
        actual = glob_context
          .query_get('select &p1 as p1, &p2 as p2, &p3 as p3', p1: 1, p2: 2, p3: 3)
        actual.must_be_instance_of AssOle::Rubify::GenericWrapper
        actual.to_s.must_match %r{Query|Запрос}
        actual.TempTablesManager.to_s.must_match %r{TempTablesManager|МенеджерВременныхТаблиц}
        value_table = actual.Execute.Unload
        value_table.Count.must_equal 1
        value_table.Get(0).p1.must_equal 1
        value_table.Get(0).p2.must_equal 2
        value_table.Get(0).p3.must_equal 3
      end

      it '#query_get with block' do
        vtable = glob_context.query_get('select &p1 as p1', p1: 1) do |q|
          q.Execute.Unload
        end
        vtable.Count.must_equal 1
        vtable.Get(0).p1.must_equal 1
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
