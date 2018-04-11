module AssOle
  module Rubify
    # Basic wrapper for 1C ole object.
    class GenericWrapper
      include Support::SendToOle

      attr_reader :ole, :ole_runtime, :owner

      # @raise ArgumentError if +ole+ or +owner+ invalid
      # @param ole [WIN32OLE] wrapped ole object
      # @param ole_runtime ole rutime which spawn +ole+ object
      # @api private
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

      def root_owner?
        owner.nil?
      end

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

