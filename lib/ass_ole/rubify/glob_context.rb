module AssOle
  module Rubify
    # FIXME: doc this global context wrapper
    class GlobContex < GenericWrapper
      # @api private
      def initialize(ole_runtime)
        super ole_runtime.ole_connector, ole_runtime
      end

      def verify!; end
    end
  end
end

