require_relative 'lib/xlsx/DSL/version'

Gem::Specification.new do |s|
  s.name        = "xlsx-DSL"
  s.version     = OpenXML::SpreadsheetML::DSL::VERSION
  s.summary     = "Excel (.xlsx) Controller, using Ruby DSL."
  s.description = "A gem using DSL for reading&writing Excel (.xlsx)."
  s.authors     = ["Cyan Yan"]
  s.email       = "cyan_cg@outlook.com"
  s.files       = `git ls-files -z`.split("\x0").reject { |f| f.match(%r{^(test|spec)/})}.select { |f| f.match /\.rb$/}
  s.homepage    = "https://rubygems.org/gems/xlsx-DSL"
  s.license     = "MIT"

  s.add_dependency 'nokogiri', '~> 1.13.10'
  s.add_dependency 'rubyzip', '~> 2.3.2'
  s.add_development_dependency 'rspec', '~> 3.12.0'
end