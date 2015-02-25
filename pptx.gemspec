require_relative 'lib/pptx/version'

Gem::Specification.new do |s|
  s.name        = 'pptx'
  s.version     = PPTX::VERSION
  s.summary     = 'PPTX'
  s.description = "Write PPTX files with slides, textboxes, and images."
  s.authors     = ["nuvyu"]
  s.email       = "contact@nuvyu.com"

  s.files       = Dir['README.md', 'lib/**/*']
  s.homepage    = 'https://github.com/nuvyu/ruby-pptx'
  s.license     = 'MIT'

  s.add_runtime_dependency 'nokogiri', '~> 1.6'
  s.add_runtime_dependency 'rubyzip', '1.1.2'
  # zipline depends on rails, so only add it as a development dependency
  s.add_development_dependency 'zipline', '0.0.9'
  s.add_development_dependency 'rspec', '~> 3.2'
end
