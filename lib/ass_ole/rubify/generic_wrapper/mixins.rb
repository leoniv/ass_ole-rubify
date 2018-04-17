module AssOle
  module Rubify
    class GenericWrapper
      # Dynamically blended mixins for {GenericWrapper} instances
      # All modules included in {Mixins} name space
      # must implements {Support::MixinInterface#blend?} which
      # returns +true+ if wrapper must be extended by this
      # mixin
      # @example (see Support::MixinInterface)
      module Mixins
        extend Support::MixinsContainer

        # (see #write)
        module Write
          def self.blend?(wr)
            wr.quack.Write?
          end

          # Overide OLE method +Write+
          # FIXME: example
          def write(*args, **attributes, &block)
            yield self if  block_given?
            _fill_attributes_(self.ole, **attributes)
            send(:Write, *args)
            self
          end
        end

        module Indexable
          extend Support::MixinsContainer

          # @api private
          def self._?(wr)
            wr.quack.Indaex? || wr.quack.UBound?
          end

          module Get
            def self.blend?(wr)
              wr.quack.Get? && Indexable._?(wr)
            end

            # FIXME: example
            def [](index)
              send(:Get, index)
            end
          end

          module Set
            def self.blend?(wr)
              wr.quack.Set? && Indexable._?(wr)
            end

            # FIXME: example
            def []=(index, value)
              send(:Set, index, value)
            end
          end
        end

        module Collection
          # FIXME
        end
      end
    end
  end
end
