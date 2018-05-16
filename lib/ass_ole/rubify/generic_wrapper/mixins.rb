module AssOle
  module Rubify
    class GenericWrapper
      # Dynamically blended mixins for {GenericWrapper} instances.
      # All modules included in {Mixins} namespace
      # must implements {Support::MixinInterface#blend?} which
      # returns +true+ if wrapper must be extended by this
      # mixin.
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

        # Mixins for 1C collections like a Array, ValueTable etc.
        # Collection is a all ole objects responds to :Count and :Get.
        # Some collections which {Indexable} with Fixnum index
        # will be ::Enumerable. But 1C Structure and Map can't
        # be Enumerable :(
        # If collection is writable, {Indexable::Set}
        # and {Add} mixins will be included.
        module Collection
          extend Support::MixinsContainer

          # Collection is a all ole objects responds to :Count and :Get
          def self._?(wr)
            wr.quack.Get? && wr.quack.Count?
          end

          # Mixin for 1C collections which indexable with Fixnum index
          # and can be Enumerable.
          # Indexable ole collection is all {Collection}
          # responds to :IndexOf && :UBound
          # @example :Enumerable
          #   a = newObject('Array') do
          #     3.times do |v|
          #       a << v + 1
          #     end
          #   end #=> GenericWrapper
          #
          #   a.map(&:to_s) #=> ["1", "2", "3"]
          #
          #   a.each_with_index do |item, index|
          #     puts "a[#{index}] = #{item}"
          #   end #=> GenericWrapper
          module Indexable
            include ::Enumerable
            extend Support::MixinsContainer

            # Indexable ole collection is all {Collection}
            # responds to :IndexOf && :UBound
            def self.blend?(wr)
              Collection._?(wr) && (wr.quack.IndexOf? || wr.quack.UBound?)
            end

            # @example
            #  a = newObject('Array') do |a|
            #    3.times do |i|
            #      a.Add(i)
            #    end
            #  end #=> GenericWrapper
            #
            #  a.each #=> GenericWrapper
            #  a.each do |item|
            #    puts item
            #  end #=> GenericWrapper
            # @return [self]
            def each
              Count().times do |i|
                yield Get(i) if block_given?
              end
              self
            end

            # Return last item
            def last
              get(Count() - 1)
            end

            # Return first item
            def first
              get(0)
            end

            # @return [Array]
            # @example
            #   a = newObject('Arra') do |a|
            #     a.Add(0)
            #     a.Add(1)
            #   end #=> GenericWrapper
            #
            #   a.to_a #=> [0, 1]
            def to_a
              map do |item|
                item
              end
            end

            # Alias for ole method Get(index) but without fail
            # when index out of range and accepts index < 0 like
            # Ruby Array
            # @example
            #   a = newObject('Array') do |a|
            #     5.times do |i|
            #       a.Add(i)
            #     end
            #   end
            #   a.Get(10) #=> WIN32OLERuntimeError index out of range
            #   a.get(10) #=> nil
            #   a[0] #=> 0
            #   a[4] #=> 4
            #   a[5] #=> nil
            #   a[-1] #=> 4
            #   a[-5] #=> 0
            #   a[-6] #=> nil
            def get(index)
              return if empty?
              return if index > Count() - 1
              if index < 0
                return if Count() < index.abs
                return Get(Count() + index)
              end
              Get(index)
            end
            alias_method :[], :get

            # Mixin for {Indexable} and writable 1C {Collection}
            # which responds to ole method :Set
            module Set
              # All ole {Indexable} {Collection} ole responds to :Set
              def self.blend?(wr)
                Indexable.blend?(wr) && wr.quack.Set?
              end

              # Alias for ole method Set(index, value) but returns value
              # @example
              #  a = newObject('Array') do |a|
              #    3.time do |i|
              #      a.Add(i)
              #    end
              #  end #=> GenericWrapper
              #
              #  a[0] #=> 0
              #  a.Set(0, 10) #=> nil
              #  a[0] = 10 #=> 10
              #  a[0] #=> 10
              #  a[10] = 0 #=> WIN32OLERuntimeError Index out of range
              # @raise [WIN32OLERuntimeError] when index out of range
              def []=(index, value)
                Set(index, value)
              end
            end
          end

          # Mixin for all collections provides {#size} method like in Ruby
          module Count
            # Suitable for all {Collection}
            def self.blend?(wr)
              wr.quack.Count?
            end

            # Alias for ole method Count()
            # @return [Fixnum]
            def size
              Count()
            end

            # True if collection Count() == 0
            def empty?
              Count() == 0
            end
          end

          # Mixin for writable 1C {Collection}
          # which responds to ole method :Add
          module Add
            # All ole {Collection} which ole responds to :Add
            def self.blend?(wr)
              Collection._?(wr) && wr.quack.Add?
            end

            # Alias for ole method Add but returns value
            # @example
            #   a = newObject('Array')
            #   a.Add(10) #=> nil
            #   a << 10 #=> 10
            def <<(value)
              Add(value)
              value
            end
          end

          # Mixin for 1C Map.
          # Unfortunately 1C Map can't be ::Enumerable
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

            # Alias for ole method Get()
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

            # Alias for ole method Insert() but returns value
            # @example
            #   m = newObject('Map')
            #   m.Insert('key', 'value') #=> nil
            #   m[:key] = 'value' #=> 'value'
            #   m['key'] = 'value' #=> 'value'
            def []=(key, value)
              Insert((key.is_a?(Symbol) ? key.to_s : key), value)
            end
          end

          # Mixin for 1C Structure.
          # Unfortunately 1C Structure can't be ::Enumerable
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

            # Alias for ole Structure property reader but doesn't fail
            # when property (or key) not exists.
            # @example
            #   s = newObject('Structure')
            #   s.key #=> WIN32OLERuntimeError
            #   s[:key] #=> nil
            #   s.Insert('key', 'value')
            #   s[:key] #=> 'value'
            #   s['key'] #=> 'value'
            def [](key)
              ole.send(key.to_s.upcase) if key?(key)
            end

            # Alias for ole method Insert() but returns value
            # @example
            #   s = newObject('Structure')
            #   s.key = 'value' #=> WIN32OLERuntimeError
            #   s.Insert('key', 'value') #=> nil
            #   s[:key] = 'value' #=> 'value'
            #   s['key'] = 'value' #=> 'value'
            def []=(key, value)
              Insert(key.to_s, value)
            end
          end
        end
      end
    end
  end
end
