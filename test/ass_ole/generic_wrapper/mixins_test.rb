require 'test_helper'
require 'ass_ole/snippets/shared'
module AssOle::RubifyTest
  describe AssOle::Rubify::GenericWrapper::Mixins do
    like_ole_runtime Runtimes::Ext

    describe 'Collection' do
      def collection_wrapper(size = 5)
        @collection_wrapper ||= AssOle::Rubify::GenericWrapper
          .new(ole_coll(size), ole_runtime_get)
      end

      def ole_coll(size)
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

        describe 'ValueTable' do
          include AssOle::Snippets::Shared::ValueTable
          include MustIncludeIndexable

          def value_table_collection(size)
            value_table :c1, :c2, :c3 do |vt|
              size.times do |i|
                vt.add c1: "c1 #{i}", c2: "c2 #{i}", c3: "c3 #{i}"
              end
            end
          end

          def ole_coll(size)
            @ole_coll ||= value_table_collection(size)
          end

          describe '#each' do
            it 'without block' do
              collection_wrapper.each.must_be_instance_of AssOle::Rubify::GenericWrapper
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
        end

        describe 'Array' do
          include AssOle::Snippets::Shared::Array
          include MustIncludeIndexable

          def array_collection(size)
            return array if size <= 0
            array(*(1 .. size).to_a)
          end

          def ole_coll(size)
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
            end

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
      end
    end
  end

end
