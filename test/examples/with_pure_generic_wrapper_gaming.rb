require 'test_helper'

module AssOle::RubifyTest
  module WithPureGenericWrapperGaming
    describe 'gaming with catalogs' do
      like_ole_runtime Runtimes::Ext
      include AssOle::Rubify

      it 'wrapp catalog manager Catalogs.Catalog1' do
        manager = rubify(cAtalogs.Catalog1)

        manager.to_s.must_equal 'CatalogManager.Catalog1'

        object = manager.CreateItem(Description: 'New object', Attribute1: 'A1 Value') do |obj|
          obj.must_be_instance_of AssOle::Rubify::GenericWrapper
          obj.Description.must_equal 'New object'
          obj.Attribute1.must_equal 'A1 Value'

          3.times do |i|
            obj.TabularSection1.Add(Attribute1: "A1 V#{i}", Attribute2: "A2 V#{i}") do |row|
              row.must_be_instance_of AssOle::Rubify::GenericWrapper
              row.Attribute1.must_match %r{A1 V#{i}}
              row.Attribute2.must_match %r{A2 V#{i}}
              row.Attribute3 = "A3 V#{i}"
            end
          end
        end

        object.must_be_instance_of AssOle::Rubify::GenericWrapper
        object.Ref.IsEmpty.must_equal true
        object.Write
        object.Ref.IsEmpty.must_equal false

        object.TabularSection1.Count.must_equal 3
        object.TabularSection1.Count.times do |index|
          row = object.TabularSection1.Get(0)
          row.must_be_instance_of AssOle::Rubify::GenericWrapper
          row.Attribute1.must_match %r{A1}
          row.Attribute2.must_match %r{A2}
          row.Attribute3.must_match %r{A3}
        end
      end

      it 'wrapp catalog manager Catalogs.Catalog2' do
        manager = rubify(cAtalogs.Catalog2)

        folder1 = manager.CreateFolder(Description: 'Folder 1')
        folder1.Write

        folder2 = manager.CreateFolder(Description: 'Folder 2', Parent: folder1.Ref)
        folder2.Write
        folder2.Parent.Description.must_equal 'Folder 1'

        folder3 = manager.CreateFolder(Description: 'Folder 3')

        ref = folder3.Parent = folder1.Ref
        ref.must_be_instance_of AssOle::Rubify::GenericWrapper

        ole_ref = folder3.Parent = folder1.ole.Ref
        ole_ref.must_be_instance_of WIN32OLE

        folder3.Parent.Description.must_equal 'Folder 1'
      end
    end

    describe 'gaiming with collections' do
      like_rubify_runtime Runtimes::Ext
      it 'Metadata.Documents' do
        mEtadata.Catalogs.map(&:to_s)
          .must_equal %w{Catalog1 Catalog2 Catalog3}
      end
    end
  end
end
