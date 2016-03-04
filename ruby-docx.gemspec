$LOAD_PATH.unshift(File.expand_path('../lib', __FILE__))
require 'docx/gem'

Gem::Specification.new do |gems|
  gems.name    = 'ruby-docx'
  gems.version = '0.0.0'
  gems.author  = 'Kevin Stenerson'
  gems.summary = 'A gem for "just" writing .docx files'
  gems.description = <<-END
Implements creating and writing of docx files.
The gem is intended to convert data from yaml/markdown/etc into a .docx
container which is compatible first with Google Drive then with Microsoft Word.
END
  gems.files = []

  gems.required_ruby_version = '>= 2.2.0'
  gems.add_development_dependency 'rubocop', '~> 0.37.0'

  # Gem builder boilerplate
  gems.email    = 'kevin@reflexionhealth.com'
  gems.homepage = 'http://reflexionhealth.com'
  gems.license  = 'Copyright (c) 2016 Reflexion Health'
end
