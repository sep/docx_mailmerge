# This file was generated by the `rspec --init` command. Conventionally, all
# specs live under a `spec` directory, which RSpec adds to the `$LOAD_PATH`.
# Require this file using `require "spec_helper"` to ensure that it is only
# loaded once.
#
# See http://rubydoc.info/gems/rspec-core/RSpec/Core/Configuration
RSpec.configure do |config|
  config.treat_symbols_as_metadata_keys_with_true_values = true
  config.run_all_when_everything_filtered = true
  config.filter_run :focus

  # Run specs in random order to surface order dependencies. If you find an
  # order dependency and want to debug it, you can fix the order by providing
  # the seed, which is printed after each run.
  #     --seed 1234
  config.order = 'random'
end

require 'rubygems'
require 'bundler'
Bundler.setup

require 'rspec'
require 'nokogiri'
require 'fileutils'
require 'docx_mailmerge'
require 'nokogiri/diff'

SPEC_BASE_PATH = Pathname.new(File.expand_path(File.dirname(__FILE__)))

RSpec::Matchers.define :be_same_xml_as do |expected|
  match do |actual|
    (Nokogiri::XML(actual).diff(Nokogiri::XML expected)).all? do |c, dummy|
      c == " "
    end
  end
  diffable
end
