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

          # Mixin for 1C Map
          # Unfortunatly 1C Map can't be ::Enumerable
          module Map
            def self.blend?(wr)
              Collection._?(wr) && wr.to_s =~ %r{Map|Соответствие}i
            end

            # @example
            #   m = newObject('Map')
            #   m.key? :key #=> false
            #   m.key? 'key' #=> false
            #   m.Insert('key', 'value')
            #   m.key? :key #=> true
            #   m.key? 'key' #=> true
            def key?(key)
              !Get((key.is_a?(Symbol) ? key.to_s : key)).nil?
            end

            # Aliase for ole method Get()
            # @example
            #   m = newObject('Structure')
            #   m['key'] #=> nil
            #   m[:key]  #=> nil
            #   m.Insert('key', 'value') #=> nil
            #   m[:key] #=> 'value'
            #   m['key'] #=> 'value'
            def [](key)
              Get((key.is_a?(Symbol) ? key.to_s : key))
            end

            # Aliase for ole method Insert()
            # @example
            #   m = newObject('Map')
            #   m.Insert('key', 'value') #=> nil
            #   m[:key] = 'value' #=> 'value'
            #   m['key'] = 'value' #=> 'value'
            def []=(key, value)
              Insert(key, value)
              value
            end
          end

          # Mixin for 1C Structure.
          # Unfortunatly 1C Structure can't be ::Enumerable
          module Structure
            def self.blend?(wr)
              wr.quack.Property? && wr.to_s =~ %r{Structure|Структура}i
            end

            # @example
            #   s = newObject('Structure')
            #   s.key? :key #=> false
            #   s.key? 'key' #=> false
            #   s.Insert('key', 'value')
            #   s.key? :key #=> true
            #   s.key? 'key' #=> true
            def key?(key)
              Property((key.is_a?(Symbol) ? key.to_s : key))
            end

            # Aliase for ole Structure propery reader
            # @example
            #   s = newObject('Structure')
            #   s.key #=> WIN32OLERuntimeError
            #   s[:key] #=> nil
            #   s.Insert('key', 'value')
            #   s[:key] #=> 'value'
            #   s['key'] #=> 'value'
            def [](key)
              ole.send(key.to_s) if key?(key)
            end

            # Aliase for ole method Insert()
            # @example
            #   s = newObject('Structure')
            #   s.key = 'value' #=> WIN32OLERuntimeError
            #   s.Insert('key', 'value') #=> nil
            #   s[:key] = 'value' #=> 'value'
            #   s['key'] = 'value' #=> 'value'
            def []=(key, value)
              Insert(key.to_s, value)
              value
            end
          end
        end
      end
    end
  end
end
