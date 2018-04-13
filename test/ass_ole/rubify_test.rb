require "test_helper"
require 'ass_ole/snippets/shared'

module AssOle::RubifyTest
  describe 'Const' do
    it 'verify version number' do
      ::AssOle::Rubify::VERSION.wont_be_nil
    end
  end

  describe AssOle::Rubify::MdManagers do

  end

  describe AssOle::Rubify do
    like_ole_runtime Runtimes::Ext
    include AssOle::Snippets::Shared::XMLSerializer

    describe '.rubify' do
      it 'when ole is WIN32OLE' do
        ole = newObject('Array')
        AssOle::Rubify::GenericWrapper.expects(:new)
          .with(ole, ole_runtime_get, nil).returns(:wrapper)
        AssOle::Rubify.rubify(ole, ole_runtime_get).must_equal :wrapper
      end

      it 'when ole is StringInternal' do
        str = ole_runtime_get.to_string_internal(newObject('array'))
        AssOle::Rubify::GenericWrapper.expects(:new)
          .with(is_a(WIN32OLE), ole_runtime_get, nil).returns(:wrapper)
        AssOle::Rubify.rubify(str, ole_runtime_get).must_equal :wrapper
      end

      it 'when ole is Xml string' do
        xml = to_xml(newObject('array'))
        AssOle::Rubify::GenericWrapper.expects(:new)
          .with(is_a(WIN32OLE), ole_runtime_get, nil).returns(:wrapper)
        AssOle::Rubify.rubify(xml, ole_runtime_get).must_equal :wrapper
      end

      it 'when ole is nil' do
        AssOle::Rubify.rubify(nil, ole_runtime_get).must_be_nil
      end

      it 'when ole is a GenericWrapper returns ole' do
        stub = Class.new(AssOle::Rubify::GenericWrapper) do
          def initialize; end
        end.new
        AssOle::Rubify.rubify(stub, nil).must_equal stub
      end

      it 'fail ArgumentError when ole is ivalid' do
        e = proc {
          AssOle::Rubify.rubify(:invalid, nil)
        }.must_raise ArgumentError
        e.message.must_match %r{Unknown ole}
      end

      it 'smoky' do
        xml = to_xml(newObject('array'))
        actual = AssOle::Rubify.rubify(xml, ole_runtime_get)
        actual.must_be_instance_of AssOle::Rubify::GenericWrapper
        actual.to_s.must_match %r{Array|Массив}
        actual.ole_runtime.must_equal ole_runtime_get
      end
    end

    describe '#rubify' do
      it 'smoky' do
        actual = Class.new do
          like_ole_runtime Runtimes::Ext
          include AssOle::Rubify
        end.new.rubify(newObject('Array'))
        actual.must_be_instance_of AssOle::Rubify::GenericWrapper
        actual.to_s.must_match %r{Array|Массив}
        actual.ole_runtime.must_equal ole_runtime_get
      end

      it 'mocked' do
        AssOle::Rubify.expects(:rubify).with(:ole, ole_runtime_get).returns(:wrapper)
        Class.new do
          like_ole_runtime Runtimes::Ext
          include AssOle::Rubify
        end.new.rubify(:ole).must_equal :wrapper
      end
    end

    it '#glob_context' do
      AssOle::Rubify.expects(:glob_context).with(ole_runtime_get)
        .returns(:GlobContex)
      inst = Class.new do
        like_ole_runtime Runtimes::Ext
        include AssOle::Rubify
      end.new
      inst.glob_context.must_equal :GlobContex
      inst.glob_context.must_equal :GlobContex, 'instace of GlobContex stored'
    end

    it '.glob_context' do
      actual = AssOle::Rubify.glob_context(ole_runtime_get)
      actual.must_be_instance_of AssOle::Rubify::GlobContex
    end

    it '.like_string_internal?' do
      AssOle::Rubify.like_string_internal?("{\n}").must_equal true
      AssOle::Rubify.like_string_internal?("0{\n}").must_equal false
    end

    it '.like_xml?' do
      AssOle::Rubify.like_xml?('  <?xml').must_equal true
      AssOle::Rubify.like_xml?('  <bla').must_equal true
      AssOle::Rubify.like_xml?('bla <?xml').must_equal false
    end

    it '.from_xml' do
      xml = to_xml(newObject('array'))
      AssOle::Rubify.from_xml(xml, ole_runtime_get).must_be_instance_of WIN32OLE
    end
  end
end
