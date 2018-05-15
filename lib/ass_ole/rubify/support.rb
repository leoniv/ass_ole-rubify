module AssOle
  module Rubify
    # Helpers mixin
    module Support
      # Helper for duck typing of +ole+
      # @example
      #   include AssOle::Rubify
      #
      #   arr = rubify(newObject('Array'))
      #
      #   # It like arr.ole_respond_to? :Get
      #   arr.quack.Get? #=> true
      class DuckTyping
        attr_reader :wrapper

        # @param wrapper [GenericWrapper]
        def initialize(wrapper)
          @wrapper = wrapper
        end

        # Invoke #ole_respond_to? symbol.to_s.gsub(/\?$/)
        # @example (see DuckTyping)
        def method_missing(symbol, *_)
          wrapper.ole_respond_to? symbol.to_s.gsub(/\?$/, '').to_sym
        end
      end

      module FillAttributes
        # @api private
        # Filling +obj+ attributes from hash +attributes+. Values of
        # +attributes+ must be valid 1C values.
        # @param obj [WIN32OLE]
        # @return obj
        def _fill_attributes_(obj, **attributes)
          attributes.each do |k, v|
            obj.send("#{k}=", v)
          end
          obj
        end
      end

      # Send message to +ole+ in {#method_missing}
      module SendToOle
        include FillAttributes

        def self.fail_must_respond_to_ole(symbol)
          fail ArgumentError, 'All included `SendToOle`'\
            ' must respond_to? `:ole`' if symbol == :ole
        end

        # Method missing handler.
        #
        # It has a two different behavior for reader and writer methods.
        #
        # Writers handling in {#\_writer_missing\_} and ignores an options
        # and the block.
        #
        # Readers handling here and sends message to the +ole+, yields wrapped
        # result of +ole+ invocation to the block
        # - Before invoke +ole+ extract real 1C values from +args+
        #   in {#\_extract_args\_}
        # - Before {#\_wrapp_ole_result\_} fills object attributes
        #   in {FillAttributes#\_fill_attributes\_} from +opts+ hash which
        #   before passed thru the {#\_extract_opts\_}
        # - Before yielding, value passes thru the {#\_wrapp_ole_result\_}
        # @yield {#_wrapp_ole_result_} wrapped +ole+ invocation result
        # @return {#_wrapp_ole_result_} wrapped +ole+ invocation result
        def method_missing(symbol, *args, **opts, &block)
          SendToOle.fail_must_respond_to_ole(symbol)
          return _writer_missing_(symbol, args[0]) if symbol.to_s =~ %r{=$}
          result = ole.send(symbol, *_extract_args_(args))
          result = _fill_attributes_(result, _extract_opts_(opts))
          result = _wrapp_ole_result_(result)
          yield result if block_given?
          result
        end

        # @api private
        # Handler for missing writer methods, send message to +ole+ and always
        # returns the same value was got in +arg+ parameter
        # @param symbol [Symbol] method name
        # @return +arg+ value
        def _writer_missing_(symbol, arg)
          ole.send(symbol, _extract_ole_(arg))
          arg
        end

        # @abstract
        # In this method will be pass value which returns from +ole+
        # invocation and you can wrap it value for goal your self
        def _wrapp_ole_result_(ole_invocation_result)
          ole_invocation_result
        end

        # @abstract
        # In this method you must extract real 1C value from +value+
        # if 1C value was wrapped in {#\_wrapp_ole_result\_}
        def _extract_ole_(value)
          value
        end

        # @api private
        # Convert values of hash +opts+ to real 1C value in {#\_extract_ole\_}
        # @param opts [Hash]
        # @return [Hash] new hash with real 1C values
        def _extract_opts_(opts)
          opts.map {|k, v| [k, _extract_ole_(v)]}.to_h
        end

        # @api private
        # Convert all values in +args+ to  real 1C value in {#\_extract_ole\_}
        # @return [Array] new array with real 1C values
        def _extract_args_(args)
          args.map {|a| _extract_ole_ a}
        end
      end

      # Mixin for {#glob_context} method
      module GlobContex
        # @return [GlobContex] instance
        # @example (see GlobContex)
        def glob_context
          @glob_context ||= AssOle::Rubify.glob_context(ole_runtime_get)
        end
      end

      # @abstract
      # Abstract interface for mixin module
      # @example (see #blend?)
      module MixinInterface
        # Returns +true+ if +wrapper+ must be extended by this mixin
        # @param wrapper [GenericWrapper]
        # @example Create new mixin
        #   module AssOle::Rubify::GenericWrapper::Mixins
        #     module Write
        #       def self.blend?(wrapper)
        #         wrapper.quack.Write?
        #       end
        #
        #       def write(*args, **options, &block)
        #         #....
        #       end
        #     end
        #   end
        def blend?(wrapper)
          fail 'Abstract method call'
        end
      end

      # TODO: doc this with example
      module MixinsContainer
        # @api private
        def blend(wr)
          constants.each do |c|
            mixin = const_get(c)
            mixin.blend(wr) if mixin.respond_to? :blend
            wr.send(:extend, mixin) if\
              mixin.respond_to?(:blend?) && mixin.blend?(wr)
          end
        end
      end
    end
  end
end

