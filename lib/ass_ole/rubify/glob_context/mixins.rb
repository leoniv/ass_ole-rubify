module AssOle
  module Rubify
    class GlobContex
      # Dynamicaly mixins for {GlobContex} instances
      # All modules included in {Mixins} name space
      # must implements {Support::MixinInterface#blend?} which
      # returns +true+ if wrapper must be extended by this
      # mixin
      # @example (see Support::MixinInterface)
      module Mixins
        extend Support::MixinsContainer

        def self._?(wr)
          wr.quack.NewObject?
        end

        # FIXME doc this
        module MixedContext
          def self.blend?(wr)
            Mixins._?(wr)
          end
        end

        # FIXME doc this
        module ServerContext
          def self.blend?(wr)
            Mixins._?(wr) && wr.quack.Metadata?
          end
        end

        # FIXME doc this
        module ClientContext
          def self.blend?(wr)
            Mixins._?(wr) && wr.quack.GetForm?
          end
        end
      end
    end
  end
end
