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
      module FillAttributes
        def _fill_attributes_(obj, **attributes)
          attributes.each do |k, v|
            obj.send("#{k}=", v)
          end
          obj
        end
        private :_fill_attributes_
      end

      module ExtractArgs
        def _extract_opts_(opts)
          opts.map {|k, v| [k, _extarct_ole_(v)]}.to_h
        end
        private :_extract_opts_

        def _extract_args_(args)
          args.map {|a| _extarct_ole_ a}
        end
        private :_extract_args_

        def _extarct_ole_(val)
          val.is_a?(GenericWrapper) ? val.ole : val
        end
        private :_extarct_ole_
      end

      # Send message to +ole+ in {#method_missing}
      module SendToOle
        include FillAttributes
        include ExtractArgs

        def self.fail_must_respond_to_ole(symbol)
          fail ArgumentError,
            'All included `SendToOle` must respond_to? `:ole`' if symbol == :ole
        end

        def method_missing(symbol, *args, **opts, &block)
          SendToOle.fail_must_respond_to_ole(symbol)

          result = _fill_attributes_(ole.send(symbol, *_extract_args_(args)),
                                     **_extract_opts_(opts))

          result = GenericWrapper.new(result, ole_runtime, self) if\
            ole_runtime.spawned? result

          yield result if block_given?

          result
        end
      end
    end

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
