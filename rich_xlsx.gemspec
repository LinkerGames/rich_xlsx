# frozen_string_literal: true
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'rich_xlsx/version'

Gem::Specification.new do |spec|
  spec.name          = "rich_xlsx"
  spec.version       = RichXlsx::VERSION
  spec.authors       = ["Felix Bünemann"]
  spec.email         = ["felix.buenemann@gmail.com"]

  spec.summary       = %q{RichXlsx is a streaming XLSX spreadsheet writer with styling support}
  spec.description   = %q{This gem allows efficient writing of data to XLSX with multiple worksheets.}
  spec.homepage      = "https://github.com/LinkerGames/rich_xlsx"
  spec.license       = "MIT"

  spec.files         = `git ls-files -z`.split("\x0").reject { |f| f.match(%r{^(test|spec|features)/}) }
  spec.bindir        = "exe"
  spec.executables   = spec.files.grep(%r{^exe/}) { |f| File.basename(f) }
  spec.require_paths = ["lib"]
  spec.required_ruby_version = ">= 2.6.0"

  spec.add_dependency "zip_kit", ">= 6.2", "< 7"

  spec.add_development_dependency "bundler", ">= 1.7"
  spec.add_development_dependency "rake"
  spec.add_development_dependency "rubyzip", ">= 1.2"
  spec.add_development_dependency "minitest"
  spec.add_development_dependency "pry"
end
