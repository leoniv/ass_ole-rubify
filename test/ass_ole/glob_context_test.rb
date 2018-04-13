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
end
