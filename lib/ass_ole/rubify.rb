require 'ass_ole'
require 'ass_ole/rubify/patches/ass_ole'
module AssOle
  # @example Usage as Mixin
  #   class Worker
  #     like_ole_runtime ExternalRuntime
  #     include AssOle::Rubify
  #
  #     def document
  #       rubify document_get
  #     end
  #   end
  #
  #   Worker.new.document.exist?
  # @example Usage as module method
  #   class Worker
  #     like_ole_runtime ExternalRuntime
  #
  #     def document
  #       AssOle::Rubify.rubify(document_get, ole_runtime_get)
  #     end
  #   end
  #
  #   Worker.new.document.exist?
  module Rubify
    require "ass_ole/rubify/version"
    require 'ass_ole/rubify/md_managers'

    # Define server context ole runtimes
    SRV_RUNTIMES = [:thick, :external]

    # Helpers mixin
    module Support
      # Send message to +ole+ in {#method_missing}
      module SendToOle
        def method_missing(symbol, *args)
          fail ArgumentError, 'All included `SendToOle`'\
            ' must respond_to? `:ole`' if symbol == :ole
          ole.send(symbol, *args)
        end
      end
    end

    # Basic wrapper for 1C ole object.
    class GenericWrapper
      include Support::SendToOle

      attr_reader :ole, :ole_runtime

      # @raise ArgumentError if +ole+ invalid
      # @param ole [WIN32OLE] wrapped ole object
      # @param ole_runtime ole rutime which spawn +ole+ object
      def initialize(ole, ole_runtime)
        @ole = ole
        @ole_runtime = ole_runtime
        yield self if block_given?
        verify!
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

      def verify!
        fail ArgumentError, "ole must be `WIN32OLE`"\
          " instance not a `#{ole.class}`" unless ole.is_a? WIN32OLE
        fail ArgumentError, 'ole must be spawned'\
          ' by ole_runtime' unless ole_runtime.spawned? ole
      end
      private :verify!
    end

    # @param ole (see GenericWrapper#initialize)
    # @return [GenericWrapper nil] wrapper for 1C +ole+ object or +nil+ if
    #  +ole.nil?+
    def rubify(ole)
      AssOle::Rubify.rubify(ole, ole_runtime_get)
    end

    # @param (see GenericWrapper#initialize)
    # @return [GenericWrapper nil] wrapper for 1C +ole+ object or +nil+ if
    #  +ole.nil?+
    def self.rubify(ole, ole_runtime)
      return ole if ole.nil?
      fail 'FIXME'
    end
  end
end
