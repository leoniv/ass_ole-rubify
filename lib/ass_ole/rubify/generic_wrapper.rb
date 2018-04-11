module AssOle
  module Rubify
    # Basic wrapper for 1C ole object.
    # All OLE objects returned +ole+ automatically wraps to {GenericWrapper}
    # in {GenericWrapper#\_wrapp_ole_result_}. Wrapper who spawns new wrapper
    # will be {GenericWrapper#owner} for new wrapper. Accordingly wrappers
    # builds hierarchical wrappers tree.
    # @example Hierarchical wrappers tree
    #   require 'ass_ole/rubify'
    #   # It's gem ass_maintainer-info_bases
    #   require 'ass_maintainer/info_bases/tmp_info_base'
    #
    #   module Runtimes
    #     module External
    #       is_ole_runtime :external
    #       run AssMaintainer::InfoBases::TmpInfoBase.new.make
    #     end
    #   end
    #
    #   module Example
    #     like_ole_runtime Runtimes::External
    #     extend AssOle::Rubify
    #
    #     md = rubify(mEtadata) #=> <AssOle::Rubify::GenericWrapper:0x2104b00c>
    #
    #     md.to_s #=> "ОбъектМетаданныхКонфигурация"
    #
    #     md.owner #=> nil
    #
    #     md.root_owner? #=> true
    #
    #     md.root_owner #=> <AssOle::Rubify::GenericWrapper:0x2104b00c>
    #
    #     md.Name #=> "Конфигурация"
    #
    #     catalogs = md.Catalogs #=> <AssOle::Rubify::GenericWrapper:0x203fb810>
    #
    #     catalogs.owner #=> <AssOle::Rubify::GenericWrapper:0x2104b00c>
    #
    #     catalogs.root_owner? #=> false
    #
    #     catalogs.root_owner #=> <AssOle::Rubify::GenericWrapper:0x2104b00c>
    #   end
    #
    #
    class GenericWrapper
      include Support::SendToOle

      # See +ole+ param of {#initialize}
      attr_reader :ole

      # See +ole_runtime+ param of {#initialize}
      attr_reader :ole_runtime

      # See +owner+ param of {#initialize}
      attr_reader :owner

      # @raise ArgumentError if +ole+ or +owner+ invalid
      # @param ole [WIN32OLE] wrapped ole object
      # @param ole_runtime ole rutime which spawn +ole+ object
      # @param owner [GenericWrapper nil] wrapper which spawn this wrapper
      #  in {#\_wrapp_ole_result\_}
      # @api private
      # @yield self
      def initialize(ole, ole_runtime, owner)
        @ole = ole
        @ole_runtime = ole_runtime
        @owner = owner
        yield self if block_given?
        verify!
      end

      # @api private
      # (see SendToOle#_wrapp_ole_result_)
      def _wrapp_ole_result_(ole_invocation_result)
        return GenericWrapper.new(ole_invocation_result, ole_runtime, self) if\
            ole_runtime.spawned? ole_invocation_result
        ole_invocation_result
      end

      # @api private
      # (see SendToOle#_extract_ole_)
      def _extract_ole_(val)
        val.is_a?(GenericWrapper) ? val.ole : val
      end

      # @return +ole_runtime.ole_connector+ ole connector to infobase
      def ole_connector
        ole_runtime.ole_connector
      end

      # @return [String] +ole+ objet's string representation
      def to_s
        ole_connector.sTring(ole).to_s
      end

      # @return (see Patches::StringInternal#to_string_internal)
      def to_string_internal
        ole_runtime.to_string_internal(ole)
      end

      # @return (see Patches::XmlTypeGet#xml_type_get)
      def xml_type
        ole_runtime.xml_type_get(ole)
      end

      # True if wrapper is on the top wrappers tree
      def root_owner?
        owner.nil?
      end

      # Returns wrapper who is on the top of wrappers tree
      # @return [GenericWrapper self]
      def root_owner
        return self if root_owner?
        owner.root_owner
      end

      def verify!
        fail ArgumentError, "ole must be `WIN32OLE`"\
          " instance not a `#{ole.class}`" unless ole.is_a? WIN32OLE
        fail ArgumentError, 'ole must be spawned'\
          ' by ole_runtime' unless ole_runtime.spawned? ole
        fail ArgumentError, 'owner must be a GenericWrapper or nil' unless\
          owner.nil? || owner.is_a?(GenericWrapper)
      end
      private :verify!
    end
  end
end

