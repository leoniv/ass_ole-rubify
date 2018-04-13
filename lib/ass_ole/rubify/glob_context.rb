module AssOle
  module Rubify
    # Wrapper for 1C global context
    # @example Access to global context
    #  module Worker
    #    like_rubify_runtime Runtimes::External
    #  end
    #
    #  Worker.glob_context #=> AssOle::Rubify::GlobContex
    #
    #  arr = Worker.newObject('Array') do |a|
    #    a.Add(1)
    #    a.Add(2)
    #  end #=> AssOle::Rubify::GenericWrapper
    #
    #  arr.Count #=> 2
    #
    #  arr.glob_context #=> AssOle::Rubify::GlobContex
    #
    #  vt = arr.glob_context
    #    .newObject('ValueTable') #=> AssOle::Rubify::GenericWrapper
    #  vt.to_s #=> "ТаблицаЗначений"
    #
    class GlobContex < GenericWrapper
      # @api private
      def initialize(ole_runtime)
        super ole_runtime.ole_connector, ole_runtime
      end

      def verify!; end
    end
  end
end

