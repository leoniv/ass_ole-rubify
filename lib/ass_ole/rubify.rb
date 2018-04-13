require 'ass_ole'

module AssOle
  # @example Usage as Mixin
  #   class Worker
  #     require 'date'
  #     like_ole_runtime ExternalRuntime
  #     include AssOle::Rubify
  #
  #     def document
  #       @document ||= document_get
  #     end
  #
  #     def document_get
  #       rubify(dOcuments).DocumentName
  #         .select(Number: '12345', Date: Date.parse('2017.01.01').to_time)
  #     end
  #   end
  #
  #   Worker.new.document.nil?
  # @example Usage as module method
  #   class Worker
  #     require 'date'
  #     like_ole_runtime ExternalRuntime
  #
  #     def document
  #       @document ||= document_get
  #     end
  #
  #     def document_get
  #       AssOle::Rubify.rubify(dOcuments, ole_runtime_get).DocumentName
  #         .select(Number: '12345', Date: Date.parse('2017.01.01').to_time)
  #     end
  #   end
  #
  #   Worker.new.document.nil?
  # @example (see GenericWrapper)
  # @example (see Dsl#like_rubify_runtime)
  # @example (see GlobContex)
  module Rubify

    # Define server context ole runtimes
    SRV_RUNTIMES = [:thick, :external]

    require 'ass_ole/rubify/patches/ass_ole'
    require 'ass_ole/rubify/version'
    require 'ass_ole/rubify/support'
    require 'ass_ole/rubify/generic_wrapper'
    require 'ass_ole/rubify/glob_context'
    require 'ass_ole/rubify/patches/core'
    require 'ass_ole/rubify/md_managers'

    include Support::GlobContex

    # Builds {GlobContex} instance
    # @param ole_runtime (see .rubify)
    # @return [GlobContex] instance
    def self.glob_context(ole_runtime)
      GlobContex.new(ole_runtime)
    end

    # @param ole [WIN32OLE GenericWrapper String] if passed +String+ expects
    # xml or +StringInternal+. If passed instance of {GenericWrapper} returs
    # same +ole+ value
    # @return [GenericWrapper nil] wrapper for 1C +ole+ object or +nil+ if
    #  +ole.nil?+
    def rubify(ole)
      AssOle::Rubify.rubify(ole, ole_runtime_get)
    end

    # @param ole [WIN32OLE GenericWrapper String] if passed +String+ expects
    # xml or +StringInternal+. If passed instance of {GenericWrapper} returs
    # same +ole+ value
    # @param ole_runtime (see GenericWrapper#initialize)
    # @return [GenericWrapper nil] wrapper for 1C +ole+ object or +nil+ if
    #  +ole.nil?+
    def self.rubify(ole, ole_runtime)
      return ole if ole.nil?
      return ole if ole.is_a? GenericWrapper
      GenericWrapper.new(ole_get(ole, ole_runtime), ole_runtime)
    end

    # @param (see .rubify)
    # @return [WIN32OLE]
    # @raise [ArgumentError] if unknown +ole+ type got
    def self.ole_get(ole, ole_runtime)
      return ole if ole.is_a? WIN32OLE
      return ole_runtime.from_string_internal(ole.to_s) if\
        like_string_internal?(ole)
      return from_xml(ole, ole_runtime) if like_xml?(ole)
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
