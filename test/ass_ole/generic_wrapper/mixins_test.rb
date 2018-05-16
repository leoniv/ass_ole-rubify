require 'test_helper'
require 'ass_ole/snippets/shared'
module AssOle::RubifyTest
  describe AssOle::Rubify::GenericWrapper::Mixins do
    like_ole_runtime Runtimes::Ext

    describe 'Collection' do
      def collection_wrapper(size = 5, **opts)
        @collection_wrapper ||= AssOle::Rubify::GenericWrapper
          .new(ole_coll(size, **opts), ole_runtime_get)
      end

      def ole_coll(size, **opts)
        fail 'Abstract method. Must return pure WIN32OLE collection'
      end

      describe 'Indexable' do
        module MustIncludeIndexable
          extend Minitest::Spec::DSL

          it 'must include Indexable' do
            collection_wrapper.singleton_class
              .include? AssOle::Rubify::GenericWrapper::Mixins::Collection::Indexable
          end
        end

        describe 'In detail tests with Array' do
          include AssOle::Snippets::Shared::Array
          include MustIncludeIndexable

          def array_collection(size)
            return array if size <= 0
            array(*(1 .. size).to_a)
          end

          def ole_coll(size, **_)
            @ole_coll ||= array_collection(size)
          end

          describe '#each' do
            it 'without block' do
              collection_wrapper.each.must_be_instance_of AssOle::Rubify::GenericWrapper
            end

            it 'with block' do
              times = 0

              collection_wrapper.each do |item|
                times += 1
                item.must_equal times
              end.must_equal collection_wrapper

              times.must_equal 5
            end
          end

          it '#each_with_index' do
            times = 0

            collection_wrapper.each_with_index do |item, index|
              times += 1
              item.must_equal times
              index.must_equal item - 1
            end.must_equal collection_wrapper

            times.must_equal 5
          end

          it '#map' do
            collection_wrapper.map(&:to_s).must_equal %w{1 2 3 4 5}
          end

          it '#size' do
            collection_wrapper.size.must_equal 5
          end

          it '#count' do
            collection_wrapper.count.must_equal 5
          end

          describe '#[index] when' do
            it 'collection is empty returns nil' do
              collection_wrapper(0).size.must_equal 0
              collection_wrapper[-1].must_be_nil
              collection_wrapper[0].must_be_nil
              collection_wrapper[1].must_be_nil
            end

            it 'index out of the range returns nil' do
              collection_wrapper.size.must_equal 5
              collection_wrapper[5].must_be_nil
              collection_wrapper[6].must_be_nil
            end

            it 'index in of the range returns item value' do
              collection_wrapper.size.must_equal 5
              collection_wrapper[0].must_equal 1
              collection_wrapper[3].must_equal 4
              collection_wrapper[4].must_equal 5
            end

            describe 'index < 0 will be reverse indexing' do
              it 'index in of the rage returns item value' do
                collection_wrapper.size.must_equal 5
                collection_wrapper[-1].must_equal 5
                collection_wrapper[-2].must_equal 4
                collection_wrapper[-3].must_equal 3
                collection_wrapper[-4].must_equal 2
                collection_wrapper[-5].must_equal 1
              end

              it 'index out of range returns nil' do
                collection_wrapper.size.must_equal 5
                collection_wrapper[-6].must_be_nil
              end
            end
          end
        end

        describe 'Reduced tests whith ValueTable' do
          include AssOle::Snippets::Shared::ValueTable
          include MustIncludeIndexable

          def value_table_collection(size)
            value_table :c1, :c2, :c3 do |vt|
              size.times do |i|
                vt.add c1: "c1 #{i}", c2: "c2 #{i}", c3: "c3 #{i}"
              end
            end
          end

          def ole_coll(size, **_)
            @ole_coll ||= value_table_collection(size)
          end

          describe '#each' do
            it 'without block' do
              collection_wrapper.each
                .must_be_instance_of AssOle::Rubify::GenericWrapper
            end

            it 'with block' do
              times = 0

              collection_wrapper.each_with_index do |row, row_index|
                [:c1, :c2, :c3].each do |column|
                  row.send(column).must_equal "#{column} #{row_index}"
                  times += 1
                end
              end.must_equal collection_wrapper

              times.must_equal 5 * 3
            end
          end

          describe '#[index] when' do
            it 'collection is empty returns nil' do
              collection_wrapper(0).size.must_equal 0
              collection_wrapper[-1].must_be_nil
              collection_wrapper[0].must_be_nil
              collection_wrapper[1].must_be_nil
            end

            it 'index out of the range returns nil' do
              collection_wrapper.size.must_equal 5
              collection_wrapper[5].must_be_nil
              collection_wrapper[6].must_be_nil
            end

            it 'index in of the range returns item value' do
              collection_wrapper.size.must_equal 5
              collection_wrapper[0].c1.must_equal "c1 0"
              collection_wrapper[3].c1.must_equal "c1 3"
              collection_wrapper[4].c1.must_equal "c1 4"
            end

            describe 'index < 0 will be reverse indexing' do
              it 'index in of the rage returns item value' do
                collection_wrapper.size.must_equal 5
                collection_wrapper[-1].c1.must_equal "c1 4"
                collection_wrapper[-2].c1.must_equal "c1 3"
                collection_wrapper[-3].c1.must_equal "c1 2"
                collection_wrapper[-4].c1.must_equal "c1 1"
                collection_wrapper[-5].c1.must_equal "c1 0"
              end

              it 'index out of range returns nil' do
                collection_wrapper.size.must_equal 5
                collection_wrapper[-6].must_be_nil
              end
            end
          end
        end

        describe 'Set' do
          include AssOle::Snippets::Shared::Array

          def ole_coll(_, **__)
            array 1, 2, 3, 4
          end

          describe '#[index]=value when' do
            it 'index in the range' do
              collection_wrapper[3].must_equal 4
              (collection_wrapper[3] = 42).must_equal 42
              collection_wrapper[3].must_equal 42
            end

            it 'index out of the range' do
              collection_wrapper[10].must_be_nil
              e = proc {
                collection_wrapper[10] = 42
              }.must_raise WIN32OLERuntimeError
              e.message.must_match %r{выходит за границы диапазона}i
            end
          end
        end
      end

      describe 'Add' do
        include AssOle::Snippets::Shared::Array

        def ole_coll(_, **__)
          array 1, 2, 3, 4
        end

        it '#<<' do
          collection_wrapper.size.must_equal 4
          (10 .. 13).to_a.each do |v|
            (collection_wrapper << v).must_equal v
          end

          collection_wrapper.size.must_equal 8

          (10 .. 13).to_a.each_with_index do |v, i|
            collection_wrapper[i + 4].must_equal v
          end
        end
      end

      describe '(Structure|Map)' do
        alias_method :structure, :collection_wrapper

        def obj_type
          fail 'Abstract must returns (Map|Structure)'
        end

        %w{Map Structure}.each do |type|
          describe type do
            def ole_coll(*_, **keys)
              obj = newObject(obj_type)
              keys.each do |k, v|
                obj.insert(k.to_s, v)
              end
              obj
            end

            define_method :obj_type do
              type
            end

            it "must include #{type}" do
              collection_wrapper.singleton_class
                .include? eval 'AssOle::Rubify::GenericWrapper::'\
                               "Mixins::Collection::#{type}"
            end

            describe '#[]' do
              it 'when has key?' do
                collection_wrapper(key: 'value')[:key].must_equal 'value'
                collection_wrapper['key'].must_equal 'value'
              end

              it 'when hasn\'t key?' do
                collection_wrapper[:key].must_be_nil
                collection_wrapper['key'].must_be_nil
              end
            end

            describe '#key?' do
              it 'when key is a Symbol' do
                collection_wrapper(key: 'value').key?(:key).must_equal true
                collection_wrapper.key?(:fakekey).must_equal false
              end

              it 'when key is String' do
                collection_wrapper(key: 'value').key?('key').must_equal true
                collection_wrapper.key?('fakekey').must_equal false
              end

              it 'when key is Object' do
                collection_wrapper.key?(newObject('Array')).must_equal false
              end
            end

            it '#[]=' do
              inst = collection_wrapper
              (inst[:key] = 'value').must_equal 'value'
              inst[:key].must_equal 'value'
              inst['key'].must_equal 'value'
              inst['key'] = 'other value'
              inst[:key].must_equal 'other value'
              inst['key'].must_equal 'other value'
              inst.Count.must_equal 1
              inst[:class] = 'class name'
              inst[:class].must_equal 'class name'
            end

            it '#size' do
              collection_wrapper(k1: nil, k2: nil, k3: nil)
                .size.must_equal 3
            end

            it '#empty?' do
              collection_wrapper.empty?.must_equal true
              collection_wrapper[:key] = 'value'
              collection_wrapper.empty?.must_equal false
            end
          end
        end
      end
    end
  end
end
