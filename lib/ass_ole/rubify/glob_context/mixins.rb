module AssOle
  module Rubify
    class GlobContex
      # Dynamically blended mixins for {GlobContex} instances
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

        # Mixins for all ole runtimes thick/thin clients
        # and external connection ole runtimes
        module MixedContext
          #@api private
          def self.blend?(wr)
            Mixins._?(wr)
          end

          # Bilds 1C Type instance by type name
          # @param type_name [String] type_name like 'String',
          #  'CatalogRef.CatName' etc
          # @return [GenericWrapper]
          # @example
          #   t = type_get 'CatalogRef.Catalog1' # => GenericWrapper
          def type_get(type_name)
            newObject('TypeDescription', type_name).Types.Get(0)
          end
        end

        # Mixins for thick client and external connection
        # runtimes
        module ServerContext
          #@api private
          def self.blend?(wr)
            Mixins._?(wr) && wr.quack.Metadata?
          end

          # Build 1C Query instance with param and TempTablesManager
          # @example
          #
          #   # Get query object
          #   query = query_get('select &p1, &p2', p1: 1, p2: 2)
          #   value_table = query.Execute.Unload
          #
          #   # Execute query without hold of query object
          #   vtable = query_get('select &p1, &p2', p1: 1, p2: 2) do |q|
          #     q.Execute.Unload
          #   end
          #
          # @param text [String] query text
          # @param tempt_manager [GenericWrapper] 1C TempTablesManager
          # @param params [Hash] qury parameters
          # @return [GenericWrapper] 1C Query instance unless &block
          #   given
          # @return &block result if block given
          # @yield [GenericWrapper] 1C Query instance
          def query_get(text, tempt_manager = nil, **params, &block)
            result = newObject('Query', text) do |q|
              q.TempTablesManager = tempt_manager ||\
                newObject('TempTablesManager')

              params.each do |k, v|
                q.SetParameter(k.to_s, v)
              end
            end
            result = yield result if block_given?
            result
          end
        end

        # Mixins for exactly thin client ole runtime
        module ClientContext
          #@api private
          def self.blend?(wr)
            Mixins._?(wr) && wr.quack.GetForm?
          end
        end
      end
    end
  end
end
