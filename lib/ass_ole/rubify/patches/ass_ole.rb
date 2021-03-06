module AssOle
  # Rubify patches for 'ass_ole'
  module Runtimes
    # Patches for 'ass_ole'
    module Patches
      # Mixins for wrapping 2 methods  +ValueFromStringInternal+ and
      # +ValueToStringInternal+ which converts 1C valuses to/from string
      # internal
      #
      # @note 1C method +ValueFromStringInternal+ and +ValueToStringInternal+
      #   defined for server runtimes only like a {Runtimes::App::External}
      #   and {Runtimes::App::Thick}.
      #
      #   For client runtime like {Runtimes::App::Thin} requires to implement
      #   {Patches::StringInternal} interface and override methods
      #   +to_string_internal+ and +from_string_internal+ in a specific
      #   +is_ole_runtime :thin+ module (see example).
      #
      # @example Implements StringInternal interface for Thin runtime
      #
      #   require 'ass_ole/rubify'
      #
      #   module ThinClientRuntime
      #     is_ole_runtime :thin
      #
      #     # In this example thin_client_helper is a common client module
      #     # defined in infobase
      #     def self.to_string_internal(value)
      #       ole_connector.thin_client_helper.ValueToStringInternal(value)
      #     end
      #
      #     # In this example thin_client_helper is a common client module
      #     # defined in infobase
      #     def self.from_string_internal(str)
      #       ole_connector.thin_client_helper.ValueFromStringInternal(str)
      #     end
      #   end
      #
      #   class OnClientWorker
      #     like_ole_runtime ThinClientRuntime
      #
      #     def to_string_internal(value)
      #       ole_runtime_get.to_string_internal(value)
      #     end
      #
      #     def from_string_internal(str)
      #       ole_runtime_get.from_string_internal(str)
      #     end
      #   end
      module StringInternal
        # NotImplementedError message for {Runtimes::App::Thin}
        def self.not_implemented
          "For client runtime like {Runtimes::App::Thin} requires override\n"\
          "methods +to_string_internal+ and +from_string_internal+ in runtime\n"\
          "module like this:\n\n"\
          "  require 'ass_ole/rubify'\n\n"\
          "  module ThinClientRuntime\n"\
          "    is_ole_runtime :thin\n\n"\
          "    # In this example thin_client_helper is a common client module\n"\
          "    # defined in infobase\n"\
          "    def self.to_string_internal(value)\n"\
          "      ole_connector.thin_client_helper.ValueToStringInternal(value)\n"\
          "    end\n\n"\
          "    # In this example thin_client_helper is a common client module\n"\
          "    # defined in infobase\n"\
          "    def self.from_string_internal(str)\n"\
          "      ole_connector.thin_client_helper.ValueFromStringInternal(str)\n"\
          "    end\n"\
          "  end\n"
        end

        # Wrapper for 1C method +ValueToStringInternal+
        # @param value any value
        # @return [String]
        def to_string_internal(value)
          ole_connector.ValueToStringInternal(value)
        end

        # Wrapper for 1C method +ValueFromStringInternal+
        # @param str [String] 1C string internalr
        # @return any value converted from +str+
        def from_string_internal(str)
          ole_connector.ValueFromStringInternal(str)
        end
      end

      # Mixin for get name of +XmlType+
      module XmlTypeGet
        # Returns name of 1C +XmlType+ for +value+. If +XmlType+ not defined for
        # +value+ returns nil.
        # @param value any value
        # @return [String nil]  +XmlType.TypeName+ of +value+
        def xml_type_get(value)
          xmlt = ole_connector.XmlTypeOf(value)
          return unless xmlt
          xmlt.TypeName
        end
      end

      # Detects that an ole object, spawned by an ole runtime
      module Spawned
        # Returns +true+ if runtime spawn +ole+ object. Always returns
        # +false+ if +ole+ is't +WIN32OLE+ instance or +ole+ is Ryby object
        # wrapped in +WIN32OLE+.
        # @param ole [WIN32OLE]
        def spawned?(ole)
          return false unless ole.is_a? WIN32OLE
          return false if ole.__ruby__?
          !ole_connector.sTring(ole).nil?
        end
      end
    end

    # Rubify patches for 'ass_ole'
    module App
      # Rubify patches for 'ass_ole'
      module External
        include Runtimes::Patches::StringInternal
        include Runtimes::Patches::XmlTypeGet
        include Runtimes::Patches::Spawned
      end

      # Rubify patches for 'ass_ole'
      module Thick
        include Runtimes::Patches::StringInternal
        include Runtimes::Patches::XmlTypeGet
        include Runtimes::Patches::Spawned
      end

      # Rubify patches for 'ass_ole'
      #
      # @example (see Runtimes::Patches::StringInternal)
      # @note (see Runtimes::Patches::StringInternal)
      module Thin
        include Runtimes::Patches::XmlTypeGet
        include Runtimes::Patches::Spawned
        # (see Runtimes::Patches::StringInternal#to_string_internal)
        # @raise [NotImplementedError]
        def to_string_internal(value)
          fail NotImplementedError, Runtimes::Patches::StringInternal
            .not_implemented
        end

        # (see Runtimes::Patches::StringInternal#from_string_internal)
        # @raise [NotImplementedError]
        def from_string_internal(str)
          fail NotImplementedError, Runtimes::Patches::StringInternal
            .not_implemented
        end
      end
    end
  end
end
