
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

      # Define class or module which transparecy invoke {GlobContex} as self
      # @example Transparency invoke {GlobContex} and wrapping all OLE to {GenericWrapper}
      #   info_base = AssMaintainer::InfoBase.new('', 'File="path"')
      #
      #   module Runtimes
      #     module External
      #       is_ole_runtime :external
      #       run infobase
      #     end
      #   end
      #
      #   module MyOleAccessor
      #     like_rubify_runtime Runtimes::External
      #   end
      #
      #   md.glob_context #=> AssOle::Rubify::GlobContex
      #   md = MyOleAccessor.Methadata #=> AssOle::Rubify::GenericWrapper
      #   md.Documents #=> AssOle::Rubify::GenericWrapper
      #   #... etc
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
