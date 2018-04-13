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

  describe 'like_rubify_runtime' do
    like_rubify_runtime Runtimes::Ext

    it 'instance has #glob_contex' do
      glob_context.must_be_instance_of AssOle::Rubify::GlobContex
    end

    describe 'all transparecy sends to #glob_context' do
      it 'smoky' do
        newObject('Array') do |a|
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
        glob_context.expects(:newObject).with('Array').returns(:Array)
        newObject('Array').must_equal 'Array'
      end
    end
  end
end
