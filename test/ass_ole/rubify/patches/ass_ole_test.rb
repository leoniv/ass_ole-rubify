require 'test_helper'

module AssOle::RubifyTest
  module Patches
    module Runtimes
      module External
        is_ole_runtime :external
        run Env::IB
      end

      module Thick
        is_ole_runtime :thick
        run Env::IB
      end

      module Thin
        module Pure
          is_ole_runtime :thin
          run Env::IB
        end

        module StrInternalImplemented
          is_ole_runtime :thin
          run Env::IB

          def self.to_string_internal(value)
            ole_connector.sTring(value)
          end

          def self.from_string_internal(str)
            ole_connector.sTring(str)
          end
        end
      end
    end

    module SharedTests
      module SrvRuntimes
        extend Minitest::Spec::DSL

        it '#to_string_internal' do
          runtimed.ole_runtime_get.to_string_internal('value')
            .must_equal '{"S","value"}'
        end

        it '#from_string_internal' do
          runtimed.ole_runtime_get.from_string_internal('{"S","value"}').must_equal 'value'
        end
      end

      module AllRuntimes
        extend Minitest::Spec::DSL

        it '#xml_type_get is string' do
          runtimed.ole_runtime_get.xml_type_get('value').must_equal 'string'
        end

        it '#xml_type_get is nil' do
          runtimed.ole_runtime_get.xml_type_get(runtimed.WebColors.Aquamarine)
            .must_be_nil
        end
      end
    end

    describe AssOle::Runtimes::App::Thick do
      include SharedTests::SrvRuntimes
      include SharedTests::AllRuntimes
      def runtimed
        @runtimed ||= Module.new do
          like_ole_runtime Runtimes::Thick
        end
      end

      it '#ole_runtime_get.type is correct' do
        runtimed.ole_runtime_get.must_equal Runtimes::Thick
      end
    end

    describe AssOle::Runtimes::App::External do
      include SharedTests::SrvRuntimes
      include SharedTests::AllRuntimes
      def runtimed
        @runtimed ||= Module.new do
          like_ole_runtime Runtimes::External
        end
      end

      it '#ole_runtime_get.type is correct' do
        runtimed.ole_runtime_get.must_equal Runtimes::External
      end
    end

    describe AssOle::Runtimes::App::Thin do
      describe 'Pure (not implements StringInternal interface)' do
        include SharedTests::AllRuntimes
        def runtimed
          @runtimed ||= Module.new do
            like_ole_runtime Runtimes::Thin::Pure
          end
        end

        it '#ole_runtime_get is correct' do
          runtimed.ole_runtime_get.must_equal Runtimes::Thin::Pure
        end

        it '#to_string_internal' do
          e = proc {
            runtimed.ole_runtime_get.to_string_internal(nil)
          }.must_raise NotImplementedError
          e.message.must_match %r{For client runtime.+requires override}
        end

        it '#from_string_internal' do
          e = proc {
            runtimed.ole_runtime_get.from_string_internal(nil)
          }.must_raise NotImplementedError
          e.message.must_match %r{For client runtime.+requires override}
        end
      end

      describe 'StrInternalImplemented' do
        include SharedTests::AllRuntimes
        def runtimed
          @runtimed ||= Module.new do
            like_ole_runtime Runtimes::Thin::StrInternalImplemented
          end
        end

        it '#ole_runtime_get is correct' do
          runtimed.ole_runtime_get.must_equal Runtimes::Thin::StrInternalImplemented
        end

        it '#to_string_internal' do
          runtimed.ole_runtime_get.to_string_internal('value').must_equal 'value'
        end

        it '#from_string_internal' do
          runtimed.ole_runtime_get.from_string_internal('value').must_equal 'value'
        end
      end
    end
  end
end
