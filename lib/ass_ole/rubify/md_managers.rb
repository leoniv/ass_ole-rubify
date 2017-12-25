module AssOle
  module Rubify
    # Provides dynamically generated mixins-helpers for access to 1C mangers
    # like as Catalogs.CatalogName, Documents.DocumentName etc. All supported
    # managers defined as modules in {MdManagers::Supported} namespase.
    #
    # Objects which includes this must includes
    # +AssOle::Runtimes::App::(Thick|External)+
    # runtime using: +like_ole_runtime+ method defined in +AssOle+ gem
    #
    # @example in scripts
    #   module PeronsWorker
    #     like_ole_runtime AccountingRuntime
    #     extend AssOle::Rubify::MdManagers::Catalogs::Persons
    #
    #     # Method try find catalog item. If fond more one item error rised.
    #     # If item not found make new item or pass exists item.
    #     def synchronize(**find_opts, &block)
    #       ref = select_object(false, **find_opts)
    #       return make_object(**find_opts, &block) unless ref
    #       result = ref.GetObject
    #       yield reusult if block_given?
    #       _write(result)
    #       result.Ref
    #     end
    #   end
    #
    #   ref = PersonWorker.synchronize(Name: 'Alice', SName: 'Cooper') do |ref|
    #     ref.Age = 99
    #   end
    #
    # @example in tests
    #
    #  describe 'Tests for Catalogs.CatalogName' do
    #    like_ole_runtime Runtimes::Ext
    #    include AssOle::Rubify::MdManagers::Catalogs[:CatalogName]
    #
    #    def catalog_item
    #      new_item Description: 'Foo' do |obj|
    #        obj.attr1 = 'bar'
    #      end
    #    end
    #
    #    it 'Item method #foo' do
    #       catalog_item.foo.must_equal 'foo'
    #    end
    #
    #    it 'ManagerModule method #bar' do
    #      catalog_manager.bar.must_equal 'bar'
    #    end
    #  end
    module MdManagers
      class TooManyObjectsFoundError < StandardError; end

      # @api private
      module Abstract
        module ObjectManager
          module GenericManager
            def object_metadata
              mEtadata.send(md_collection).send(self.MD_NAME)
            end

            def objects_manager
              fail "Manager #{md_collection}.#{self.MD_NAME} not found" unless\
                send(md_collection).ole_respond_to? self.MD_NAME
              send(md_collection).send(self.MD_NAME)
            end
          end

          include GenericManager
          include AssOle::Snippets::Shared::Query
          # Symbol like a :Catalogs, :Constants
          def md_collection
            fail 'Abstract method call'
          end

          # Abstract method must be redefined in module
          # For example catatlogs makes with method :CreateItem
          # but for Documents uses method :CreateDocument
          def _new_object_method
            fail 'Abstract method call'
          end
          private :_new_object_method

          def _new_object
            objects_manager.send(_new_object_method)
          end
          private :_new_object

          # Make new object. Only make and not write
          # @return [WIN32OLE] object not ref!
          def new_object(**attributes, &block)
            obj = _fill_attr(_new_object, **attributes)
            yield obj if block_given?
            obj
          end

          def _fill_attr(obj, **attributes)
            attributes.each do |k, v|
              obj.send("#{k}=", v)
            end
            obj
          end
          private :_fill_attr

          # Make new object and write them
          # @return [WIN32OLE] object not ref!
          def make_object(**attributes, &block)
            obj = new_object(**attributes, &block)
            _write(obj)
            obj
          end

          # If for write real object uses not :Write method it must be
          # redefined in module
          def _write(obj)
            obj.Write()
          end

          # Find multiple objects returns +nil+ or +Array+ of found
          def find_objects(*args, **options, &block)
            fail 'Abstract method call'
          end

          # Find simple object returns +nil+ or found. Riase error if multiple
          # objects found.
          def select_object(*args, **options, &block)
            result = find_objects(*args, **options, &block)
            return result if result.nil?
            return result[0] if result.size <= 1
            fail(TooManyObjectsFoundError,
                 "Find param: #{args}, #{options}") if result.size > 1
          end

          def qtext_condition_(**options)
            qtext = ''
            options.keys.each_with_index do |key, index|
              op = index != 0 ? 'and' : ''
              qtext << "#{op} ref.#{key} = &#{key}\n"
            end
            qtext
          end
          private :qtext_condition_

          def ole_to_arr_(ole_arr)
            r = []
            ole_arr.each do |i|
              r << i
            end
            r
          end
          private :ole_to_arr_
        end

        module GroupedObject
          # Make new object grouo. Only make and not write
          # @return [WIN32OLE] object not ref!
          def new_group(**attributes, &block)
            obj = _fill_attr(_new_group, **attributes)
            yield obj if block_given?
            obj
          end

          def _new_group
            objects_manager.send(:CreateFolder)
          end
          private :_new_group

          # Make new object group and write them
          # @return [WIN32OLE] object not ref!
          def make_group(**attributes, &block)
            obj = new_group(**attributes, &block)
            _write(obj)
            obj
          end

          def __folding__?
            def folding_group?
              return true unless object_metadata.ole_respond_to? :HierarchyType
              sTring(object_metadata.HierarchyType) =~\
              %r{(ИерархияГруппИЭлементов|HierarchyFoldersAndItems)}
            end
            object_metadata.Hierarchical && folding_group?
          end
          private :__folding__?
        end

        module RegisterManager
          include AssOle::Snippets::Shared::Structure
          include AssOle::Snippets::Shared::Array
          include Abstract::ObjectManager::GenericManager
          require 'date'

          alias_method :register_manager, :objects_manager
          alias_method :register_metadata, :object_metadata

          def register_record_key
            r = [:Recorder, :LineNumber] + register_dimensions
            r.unshift :Period if period_key?
            r.select do |k|
              record_key.ole_respond_to? k
            end
          end

          def information_register?
            md.InformationRegisterPeriodicity if\
              md.ole_respond_to? :InformationRegisterPeriodicity
          end

          def periodicity
            returt unless information_register?
            ole_connector.sTring md.InformationRegisterPeriodicity
          end

          # Returns false only for Nonperiodical InformationRegister
          def period_key?
            return false if periodicity
              .to_s =~ %r{(Непериодический|Nonperiodical)}i
            true
          end

          def register_dimensions
            r = []
            register_metadata.Dimensions.each do |d|
              r << d.Name.to_sym
            end
            r
          end

          def record_set(**options, &block)
            rs = register_manager.CreateRecordSet
            options.each do |k, v|
              rs.Filter.send(k).Set(v, true)
            end
            yield rs if block_given?
            rs
          end

          def record_key(options = {})
            register_manager.CreateRecordKey(structure(**options))
          end
        end
      end

      # @todo add DocumentManger, TaskManger etc.
      module Supported
        #@api private
        def self.md_collections
          constants.map do |const_name|
            result = const_get(const_name)

            result.send :define_method, :md_collection do
              const_name
            end

            result.send :define_singleton_method, :md_collection do
              const_name
            end

            result
          end
        end

        module Catalogs
          include Abstract::ObjectManager
          include Abstract::GroupedObject

          def _new_object_method
            :CreateItem
          end

          alias_method :catalog_manager, :objects_manager
          alias_method :catalog_metadata, :object_metadata
          alias_method :new_item, :new_object
          alias_method :make_item, :make_object
          alias_method :new_folder, :new_group
          alias_method :make_folder, :make_group

          def find_objects(is_folder = false, **options, &block)
            qtext = "Select T.ref as ref from\n"\
              " Catalog.#{self.MD_NAME} as T\n where \n"

            options[:IsFolder] = is_folder if __folding__?

            qtext << qtext_condition_(**options)

            arr = ole_to_arr_(query(qtext, **options)
              .Execute.Unload.UnloadColumn('ref'))
            return if arr.empty?
            return yield arr if block_given?
            arr
          end
        end

        module Documents
          include Abstract::ObjectManager

          def _new_object_method
            :CreateDocument
          end

          alias_method :document_manager, :objects_manager
          alias_method :document_metadata, :object_metadata
          alias_method :new_doc, :new_object
          alias_method :make_doc, :make_object

          def post_(doc)
            doc.write(documentWriteMode.Posting)
          end

          def undo_post_(doc)
            doc.write(documentWriteMode.UndoPosting)
          end

          # @todo find document per +date_period+ in
          #  period Metadata.NumberPeriodicity
          def find_objects(**options, &block)
            qtext = "Select T.ref as ref from\n"\
              " Document.#{self.MD_NAME} as T\n where \n"

            day = options.delete(:Дата) || options.delete(:Date)

            if day
              qtext << "BEGINOFPERIOD(T.Date, DAY) = BEGINOFPERIOD(&Date, DAY)\n"
              qtext << "AND \n" if options.size > 0
              qtext << qtext_condition_(**options)
              options[:Date] = day
            else
              qtext << qtext_condition_(**options)
            end

            arr = ole_to_arr_(query(qtext, **options)
              .Execute.Unload.UnloadColumn('ref'))
            return if arr.empty?
            return yield arr if block_given?
            arr
          end
        end

        module Constants
          include Abstract::ObjectManager::GenericManager

          alias_method :constant_manager, :objects_manager
          alias_method :constant_metadata, :object_metadata

          def constant_value_manager
            constant_manager.CreateValueManager
          end
        end

        module InformationRegisters
          include Abstract::RegisterManager

          def record_manager(**options, &block)
            rm = register_manager.CreateRecordManager
            options.each do |k, v|
              rm.send("#{k}=", v)
            end
            yield rm if block_given?
            rm
          end
        end

        module ExternalDataProcessors
          def external_object
            send(md_collection).Create(self.MD_NAME)
          end

          def ole_version
             Gem::Version.new ole_connector.NewObject('SystemInfo').AppVersion
          end

          # @todo Fuckin 1C. Refactoring require
          def external_connect(path, safe_mode = false)
            external_object_path = AssLauncher::Support::Platforms
              .path(path.to_s).realpath

            dd = ole_connector
              .newObject('BinaryData', external_object_path.win_string)

            link = ole_connector.putToTempStorage(dd)

            if ole_version > Gem::Version.new('8.3.9.2033')
              aod = ole_connector
                .newObject 'UnsafeOperationProtectionDescription'
              aod.UnsafeOperationWarnings = false
              ole_connector.send(md_collection)
                .connect(link, self.MD_NAME, safe_mode, aod)
            else
              ole_connector.send(md_collection)
                .connect(link, self.MD_NAME, safe_mode)
            end
          end
        end

        module ExternalReports
          include ExternalDataProcessors
        end

        module ChartsOfAccounts
          include Abstract::ObjectManager
          include Abstract::GroupedObject

          def _new_object_method
            :CreateAccount
          end

          alias_method :account_manager, :objects_manager
          alias_method :account_metadata, :object_metadata
          alias_method :new_account, :new_object
          alias_method :make_account, :make_object

          def find_objects(**options, &block)
            qtext = "Select T.ref as ref from\n"\
              " ChartOfAccounts.#{self.MD_NAME} as T\n where \n"

            qtext << qtext_condition_(**options)

            arr = ole_to_arr_(query(qtext, **options)
              .Execute.Unload.UnloadColumn('ref'))
            return if arr.empty?
            return yield arr if block_given?
            arr
          end
        end
      end

      # @api private
      module ManagersHolder
        def manager_holder?
          true
        end

        def content
          @content ||= {}
        end

        def [](md_name)
          r = content_get(md_name.to_s)
          fail ArgumentError,
            "#{md_collection_name}[:#{md_name}] not found" unless r
          r
        end

        def content_get(md_name)
          return content[md_name] if content[md_name]

          md_object = md_object_get(md_name)
          content[md_object.Name] = build_manager(md_object, abstract_manager)
          content[md_name]
        end

        def md_object_get(md_name)
          Class.new do
            def initialize(md_name)
              @md_name = md_name
            end

            def Name
              @md_name
            end
          end.new(md_name)
        end

        def build_manager(md_object, abstract_manager)
          r = Module.new do
            include abstract_manager

            define_method :MD_NAME do
              md_object.Name
            end
          end
        end

        def method_missing(method, *args)
          content_get(method)
        end

        attr_reader :abstract_manager
        attr_reader :md_collection_name
      end

      # @api private
      def self.build_manager_holder(abstract_manager)
        Module.new do
          extend ManagersHolder
          @abstract_manager = abstract_manager
          @md_collection_name = abstract_manager.md_collection
        end
      end

      # @api private
      def self.init
        Supported.md_collections.each do |abstract_manager|
          const_set abstract_manager.md_collection,
            build_manager_holder(abstract_manager)
        end
      end

      init
    end
  end
end
