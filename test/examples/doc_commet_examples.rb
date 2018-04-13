require 'test_helper'

module AssOle::RubifyTest
  describe 'Doc comment examples' do
    it 'Access to global context example' do
      worker = Module.new do
        like_rubify_runtime Runtimes::Ext
      end

      worker.glob_context.must_be_instance_of AssOle::Rubify::GlobContex

      arr = worker.newObject('Array') do |a|
        a.Add(1)
        a.Add(2)
      end

      arr.must_be_instance_of AssOle::Rubify::GenericWrapper

      arr.Count.must_equal 2

      arr.glob_context.must_be_instance_of AssOle::Rubify::GlobContex

      vt = arr.glob_context.newObject('ValueTable')

      vt.must_be_instance_of AssOle::Rubify::GenericWrapper

      vt.to_s.must_match %r{ТаблицаЗначений|ValueTable}
    end
  end
end
