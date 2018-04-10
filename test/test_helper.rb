$LOAD_PATH.unshift File.expand_path("../../lib", __FILE__)
require 'ass_ole/rubify'
require 'ass_maintainer/info_base'

module AssOle::RubifyTest
  module Env
    extend AssLauncher::Api
    PLATFORM_REQUIRE = '~> 8.3.10.0'
    TMP_DIR = Dir.tmpdir
    FIXT_DIR = File.expand_path("../fixtures", __FILE__)
    IB_TEMPLATE = File.join(FIXT_DIR, 'ib.cf')

    def self.make_ib(name, tmplt)
      AssMaintainer::InfoBase
        .new(name, cs_file(file: File.join(TMP_DIR, name)), false,
             platform_require: PLATFORM_REQUIRE,
             after_make: proc {|ib| ib.cfg.load(tmplt) && ib.db_cfg.update}).rebuild! :yes
    end

    IB = make_ib('ass_ole-rubify_test.ib', IB_TEMPLATE)
  end

  module Runtimes
    module Ext
      is_ole_runtime :external
      run Env::IB
    end

    module Thin
      is_ole_runtime :thin
      run Env::IB
      ole_connector.Visible = false
    end
  end
end

require "minitest/autorun"
require 'mocha/minitest'
