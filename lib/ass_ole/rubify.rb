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
