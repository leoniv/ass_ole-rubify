module AssOle
  module Rubify
    # Helpers mixin
    module Support
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

        # Send message to +ole+ yiels wrapped returned value to block.
        # - Before invoke +ole+ extract real 1C values from +args+
        #   in {#\_extract_args\_}
        # - Before {#\_wrapp_ole_result\_} fills object attributes
        #   in {FillAttributes#\_fill_attributes\_} from +opts+ hash which bebore
        #   passed thow #{#\_extract_opts\_}
        # - Before yiels invoke {#\_wrapp_ole_result\_}.
        # @yield {#_wrapp_ole_result_} wrapped +ole+ invocation result
        # @return {#_wrapp_ole_result_} wrapped +ole+ invocation result
        def method_missing(symbol, *args, **opts, &block)
          SendToOle.fail_must_respond_to_ole(symbol)
          result = ole.send(symbol, *_extract_args_(args))
          result = _fill_attributes_(result, _extract_opts_(opts))
          result = _wrapp_ole_result_(result)
          yield result if block_given?
          result
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
    end
  end
end

