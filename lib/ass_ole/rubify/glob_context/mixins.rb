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

        # FIXME doc this
        module MixedContext
          def self.blend?(wr)
            Mixins._?(wr)
          end

          # FIXME doc this
          def Array(*args)
            fail 'FIXME'
          end
          alias_method :array, :Array

          # FIXME doc this
          def Structure(**opts)
            fail 'FIXME'
          end
          alias_method :structure, :Structure

          # FIXME doc this
          def Map(**opts)
            fail 'FIXME'
          end
          alias_method :map, :Map

          # FIXME doc this
          def Type(type_name)
            newObject('TypeDescription', type_name).Types.Get(0)
          end
          alias_method :type, :Type
        end

        # FIXME doc this
        module ServerContext
          def self.blend?(wr)
            Mixins._?(wr) && wr.quack.Metadata?
          end

          # FIXME doc this
          def ValueTable(**columns)
            fail 'FIXME'
          end
          alias_method :value_table, :ValueTable

          # FIXME doc this
          def Query(text, temp_tables_manager, **param)
            fail 'FIXME'
          end
          alias_method :query, :Query
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
