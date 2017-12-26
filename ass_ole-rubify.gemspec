# coding: utf-8
lib = File.expand_path("../lib", __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require "ass_ole/rubify/version"

Gem::Specification.new do |spec|
  spec.name          = "ass_ole-rubify"
  spec.version       = AssOle::Rubify::VERSION
  spec.authors       = ["Leonid Vlasov"]
  spec.email         = ["leoniv.vlasov@gmail.com"]

  spec.summary       = %q{Wrappers for 1C WIN32OLE objects}
  spec.description   = %q{It's a part of https://github.com/leoniv/ass_ole stack. Gem provides wrappers which make Ruby scripting over 1C WIN32OLE objects easly and shortly}
  spec.homepage      = "https://github.com/leoniv/ass_ole-rubify"

  spec.files         = `git ls-files -z`.split("\x0").reject do |f|
    f.match(%r{^(test|spec|features)/})
  end
  spec.bindir        = "exe"
  spec.executables   = spec.files.grep(%r{^exe/}) { |f| File.basename(f) }
  spec.require_paths = ["lib"]

  spec.add_dependency 'ass_ole', '~> 0.3'
  spec.add_dependency 'ass_ole-snippets-shared'

  spec.add_development_dependency "bundler", "~> 1.15"
  spec.add_development_dependency "rake", "~> 10.0"
  spec.add_development_dependency "minitest", "~> 5.0"
  spec.add_development_dependency "pry"
  spec.add_development_dependency "simplecov"
  spec.add_development_dependency "mocha"
  spec.add_development_dependency "ass_maintainer-info_base"
end
