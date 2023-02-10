Gem::Specification.new do |s|
  s.name        = "xlsx-DSL"
  s.version     = "0.0.0"
  s.summary     = "Excel (.xlsx) Controller, using Ruby DSL."
  s.description = "A gem using DSL for reading&writing Excel (.xlsx)."
  s.authors     = ["Cyan Yan"]
  s.email       = "cyan_cg@outlook.com"
  s.files       = `git ls-files -z`.split("\x0").reject { |f| f.match(%r{^(test|spec)/})}.select { |f| f.match /\.rb$/}
  s.homepage    = "https://rubygems.org/gems/xlsx-DSL"
  s.license     = "MIT"
end