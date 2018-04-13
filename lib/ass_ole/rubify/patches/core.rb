
module AssOle
  module Rubify
    module Dsl
      # @api private
      module Support
        module SendToGlobContext
          include Rubify::Support::GlobContex

          def method_missing(symbol, *args, **opts, &block)
            AssOle::Snippets.fail_if_bad_context(self)
            glob_context.send(symbol, *args, **opts, &block) if ole_connector
          end
        end
      end

      # FIXME: doc this
      def like_rubify_runtime(runtime)
        like_ole_runtime runtime
        case self
        when Class then
          include Support::SendToGlobContext
        else
          extend Support::SendToGlobContext
        end
      end
    end
  end
end

class Module
  include AssOle::Rubify::Dsl
end
