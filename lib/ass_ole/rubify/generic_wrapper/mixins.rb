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

#FIXME         module Indexable
#FIXME           extend Support::MixinsContainer
#FIXME
#FIXME           # @api private
#FIXME           def self._?(wr)
#FIXME             wr.quack.Indaex? || wr.quack.UBound?
#FIXME           end
#FIXME
#FIXME           module Get
#FIXME             def self.blend?(wr)
#FIXME               wr.quack.Get? && Indexable._?(wr)
#FIXME             end
#FIXME
#FIXME             # FIXME: example
#FIXME             def [](index)
#FIXME               send(:Get, index)
#FIXME             end
#FIXME           end
#FIXME
#FIXME           module Set
#FIXME             def self.blend?(wr)
#FIXME               wr.quack.Set? && Indexable._?(wr)
#FIXME             end
#FIXME
#FIXME             # FIXME: example
#FIXME             def []=(index, value)
#FIXME               send(:Set, index, value)
#FIXME             end
#FIXME           end
#FIXME         end

        # FIXME
        module Collection
          extend Support::MixinsContainer

          def self._?(wr)
            wr.quack.Get? && wr.quack.Count?
          end

          # FIXME
          module Indexable
            include ::Enumerable
            extend Support::MixinsContainer

            def self.blend?(wr)
              Collection._?(wr) && (wr.quack.IndexOf? || wr.quack.UBound?)
            end

            def each
              Count().times do |i|
                yield Get(i) if block_given?
              end
              self
            end

            def size
              Count()
            end

            def empty?
              size == 0
            end

            def last
              get(size - 1)
            end

            def first
              get(0)
            end

            def to_a
              map do |item|
                item
              end
            end

            def get(index)
              return if empty?
              return if index > size - 1
              if index < 0
                return if size < index.abs
                return Get(size + index)
              end
              Get(index)
            end
            alias_method :[], :get

            # FIXME
            module Set
              def self.blend?(wr)
                Indexable.blend?(wr) && wr.quack.Set?
              end

              def []=(index, value)
                Set(index, value)
              end
            end
          end
        end
      end
    end
  end
end
