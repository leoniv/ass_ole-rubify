module AssOle
  module Rubify
    class GenericWrapper
      # Dynamicaly mixins for {GenericWrapper} instances
      # All modules included in {Mixins} name space
      # must implements {MixinInterface#bland?} which
      # returns +true+ if wrapper must be extended by this
      # mixin
      # @example (see MixinInterface)
      module Mixins
        # @abstract
        # Abstract interface for {GenericWrapper} mixin module
        # @example (see #bland?)
        module MixinInterface
          # Returns +true+ if +wrapper+ must be extended by this mixin
          # @param wrapper [GenericWrapper]
          # @example Writes mixin example
          #   module AssOle::Rubify::GenericWrapper::Mixins
          #     module Write
          #       def self.bland?(wrapper)
          #         wrapper.quack.Write?
          #       end
          #
          #       def write(*args, **options, &block)
          #         #....
          #       end
          #     end
          #   end
          def bland?(wrapper)
            fail 'Abstract method call'
          end
        end

        # (see #write)
        module Write
          def self.bland?(wr)
            wr.quack.Write?
          end

          # Overide OLE method +Write+
          def write(*args, **attributes, &block)
            yield self if  block_given?
            _fill_attributes_(self.ole, **attributes)
            send(:Write, *args)
            self
          end
        end

      end
    end
  end
end
