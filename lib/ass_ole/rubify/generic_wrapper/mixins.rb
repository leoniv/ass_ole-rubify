module AssOle
  module Rubify
    class GenericWrapper
      # Dynamicaly mixins for {GenericWrapper} instances
      # All modules included in {Mixins} name space
      # must implements {Helpers::MixinInterface#bland?} which
      # returns +true+ if wrapper must be extended by this
      # mixin
      # @example (see Helpers::MixinInterface)
      module Mixins
        module Helpers
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

          # TODO: doc this with example
          module MixinsContainer
            # @api private
            def bland(wr)
              constants.each do |c|
                mixin = const_get(c)
                mixin.bland(wr) if mixin.respond_to? :bland
                wr.send(:extend, mixin) if\
                  mixin.respond_to?(:bland?) && mixin.bland?(wr)
              end
            end
          end
        end

        extend Helpers::MixinsContainer

        # (see #write)
        module Write
          def self.bland?(wr)
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
          extend Helpers::MixinsContainer

          # @api private
          def self._?(wr)
            wr.quack.Indaex? || wr.quack.UBound?
          end

          module Get
            def self.bland?(wr)
              wr.quack.Get? && Indexable._?(wr)
            end

            # FIXME: example
            def [](index)
              send(:Get, index)
            end
          end

          module Set
            def self.bland?(wr)
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
