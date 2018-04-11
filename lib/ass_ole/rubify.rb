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

    # Define server context ole runtimes
    SRV_RUNTIMES = [:thick, :external]

    require 'ass_ole/rubify/version'
    require 'ass_ole/rubify/support'
    require 'ass_ole/rubify/generic_wrapper'
    require 'ass_ole/rubify/md_managers'

    # @param ole (see .rubify)
    # @return [GenericWrapper nil] wrapper for 1C +ole+ object or +nil+ if
    #  +ole.nil?+
    def rubify(ole)
      AssOle::Rubify.rubify(ole, ole_runtime_get)
    end

    # @param ole [WIN32OLE String] if passed +String+ expects xml or
    #  +StringInternal+
    # @param ole_runtime (see GenericWrapper#initialize)
    # @return [GenericWrapper nil] wrapper for 1C +ole+ object or +nil+ if
    #  +ole.nil?+
    def self.rubify(ole, ole_runtime)
      return ole if ole.nil?
      return GenericWrapper.new(ole, ole_runtime, nil) if ole.is_a? WIN32OLE
      return GenericWrapper.new(ole_runtime.from_string_internal(ole.to_s),
                                ole_runtime, nil) if like_string_internal?(ole)
      return GenericWrapper.new(from_xml(ole, ole_runtime),
                                ole_runtime, nil) if like_xml?(ole)
      fail ArgumentError, "Unknown ole: `#{ole}`"
    end

    # @api private
    def self.like_string_internal?(str)
      !(str.to_s =~ %r(\A\{.+\}\z)m).nil?
    end

    # @api private
    def self.like_xml?(str)
      !(str.to_s =~ %r{\A\s*<(\?xml|\S+)}i).nil?
    end

    # @api private
    def self.from_xml(str, ole_runtime)
      Module.new do
        like_ole_runtime ole_runtime
        extend AssOle::Snippets::Shared::XMLSerializer
      end.from_xml(str.to_s)
    end
  end
end
