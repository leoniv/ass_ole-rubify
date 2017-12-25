require 'ass_ole'

module AssOle
  # @example
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
  module Rubify
    require "ass_ole/rubify/version"
    require 'ass_ole/rubify/md_managers'

    # Helpers mixin
    module Support
      # Send message to +ole+ in {#method_missing}
      module SendToOle
        def method_missing(symbol, *args)
          ole.send(symbol, *args)
        end
      end
    end

    class GenericWrapper
      include Support::SendToOle
      attr_reader :ole, :ole_runtime
      def initialize(ole, ole_runtime)
        @ole = ole
        @ole_runtime = ole_runtime
      end

      def ole_connector
        ole_runtime.ole_connector
      end
    end

    def rubify(ole)
      AssOle::Rubify.rubify(ole, ole_runtime_get)
    end

    def self.rubify(ole, ole_runtime)
      fail 'FIXME'
    end
  end
end
