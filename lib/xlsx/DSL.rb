rbs_path = File.expand_path('DSL', File.dirname(__FILE__))
for file in Dir.entries(rbs_path)
  require(File.expand_path($1, rbs_path)) if /(.*)\.rb$/.match(file)
end