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
        require 'ass_ole/snippets/shared'
        extend Support::MixinsContainer

        def self._?(wr)
          wr.quack.NewObject?
        end

        # FIXME doc this
        module MixedContext
          def self.blend?(wr)
            Mixins._?(wr)
          end

          # FIXME doc this
          include AssOle::Snippets::Shared::Array
          alias_method :Array, :array
          # FIXME doc this
          include AssOle::Snippets::Shared::Structure
          alias_method :Structure, :structure
          # FIXME doc this
          include AssOle::Snippets::Shared::Map
          alias_method :Map, :map

          # FIXME doc this
          def Type(type_name)
            newObject('TypeDescription', type_name).Types.Get(0)
          end
        end

        # FIXME doc this
        module ServerContext
          def self.blend?(wr)
            Mixins._?(wr) && wr.quack.Metadata?
          end

          # FIXME doc this
          include AssOle::Snippets::Shared::ValueTable
          alias_method :ValueTable, :value_table
          # FIXME doc this
          include AssOle::Snippets::Shared::Query
          alias_method :Query, :query
          alias_method :TempTablesManager, :temp_tables_manager
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
